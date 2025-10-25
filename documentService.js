const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
let chokidar = null;
try { chokidar = require('chokidar'); } catch (_err) { chokidar = null; }
const ExcelJS = require('exceljs');
const db = require('./db');
const msal = require('@azure/msal-node');
const ahmenCosting = require('./ahmenCosting');
const { BrowserWindow } = require('electron');
const SETTINGS_PATH = path.join(__dirname, 'settings.json');
function readSettings() {
  try {
    const raw = fs.readFileSync(SETTINGS_PATH, 'utf-8');
    return JSON.parse(raw);
  } catch (err) {
    return {};
  }
}
function writeSettings(next) {
  try {
    fs.writeFileSync(SETTINGS_PATH, JSON.stringify(next, null, 2), 'utf-8');
    return true;
  } catch (_) { return false; }
}
const settings = readSettings();
const os = require('os');

const INVALID_FILENAME_CHARS = /[\\/:*?"<>|]/g;
const TEMPLATE_BINDING_KEY = 'ahmen_excel';
const PLACEHOLDER_PATTERN = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;

const PROCESS_TYPE = typeof process !== 'undefined' ? process.type : undefined;
const IS_MAIN_PROCESS = PROCESS_TYPE === 'browser' || PROCESS_TYPE === undefined;
const SCHEDULED_EMAIL_POLL_INTERVAL_MS = 60 * 1000;
const SCHEDULED_EMAIL_BATCH_SIZE = 10;

let scheduledMailWorkerStarted = false;
let scheduledMailWorkerExecuting = false;
let ElectronBrowserWindow = null;

function normalizeTokenKey(value) {
  if (!value) return '';
  return String(value)
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function formatCurrencyGBP(value) {
  const num = Number(value);
  const v = Number.isFinite(num) ? num : 0;
  try {
    return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(v);
  } catch (_) {
    return `£${v.toFixed(2)}`;
  }
}

function escapeHtml(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildItemsTableHtml(items = []) {
  const safe = Array.isArray(items) ? items : [];
  const rows = safe.map(it => {
    const qty = it.quantity != null && Number.isFinite(it.quantity) ? String(it.quantity) : '';
    const rate = it.rate != null && Number.isFinite(it.rate) ? formatCurrencyGBP(it.rate) : '';
    const total = it.amount != null && Number.isFinite(it.amount) ? formatCurrencyGBP(it.amount) : '';
    return `<tr><td>${escapeHtml(it.description || '')}</td><td style="text-align:right">${qty}</td><td>${escapeHtml(it.unit || '')}</td><td style="text-align:right">${rate}</td><td style="text-align:right">${total}</td></tr>`;
  }).join('');
  return `<table style="width:100%;border-collapse:collapse;margin-top:12px"><thead><tr style="background:#f9fafb"><th style="text-align:left;padding:8px;border-bottom:1px solid #e5e7eb">Description</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:80px">Qty</th><th style="text-align:left;padding:8px;border-bottom:1px solid #e5e7eb;width:80px">Unit</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:120px">Rate</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:120px">Line Total</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function renderTemplateHtml(template, tokens) {
  if (!template) return '';
  let html = String(template);
  Object.entries(tokens || {}).forEach(([key, value]) => {
    const re = new RegExp(`{{\\s*${key}\\s*}}`, 'gi');
    html = html.replace(re, value != null ? String(value) : '');
  });
  return html;
}

// Normalize outgoing email HTML to consistent typography across templates and signatures
function normalizeEmailHtml(input, opts = {}) {
  try {
    const baseFamily = String(opts.baseFamily || 'Arial, Helvetica, sans-serif');
    const baseSize = String(opts.baseSize || '10pt');
    const baseLineHeight = String(opts.baseLineHeight || '1.45');
    let html = String(input || '');
    if (!html) return '';
    // Strip inline font-size declarations to avoid mixed sizes
    html = html.replace(/font-size\s*:\s*[^;"']+;?/gi, '');
    // Optionally strip conflicting font-family declarations (keep link/icon fonts intact by not being overly aggressive)
    // html = html.replace(/font-family\s*:\s*[^;"']+;?/gi, '');
    // Remove empty style attributes (e.g., style=" ") left behind
    html = html.replace(/\sstyle=\"\s*\"/gi, '');
    // Wrap in a base container to set default family/size/line-height
    // Avoid duplicating wrapper if already present
    const trimmed = html.trim();
    const wrapperOpen = `<div style="font-family:${baseFamily};font-size:${baseSize};line-height:${baseLineHeight};color:#3c3c3b;">`;
    const wrapperClose = '</div>';
    if (!/^<div\b[^>]*>/.test(trimmed) || !/font-size:\s*\d/i.test(trimmed)) {
      return `${wrapperOpen}${trimmed}${wrapperClose}`;
    }
    return trimmed;
  } catch (_err) {
    return String(input || '');
  }
}

// Variant that preserves the signature region (wrapped by composer with comment markers)
function normalizeEmailHtmlPreserveSignature(input, opts = {}) {
  try {
    const SIG_START = '<!--__IM_SIG_START__-->';
    const SIG_END = '<!--__IM_SIG_END__-->';
    const baseFamily = String(opts.baseFamily || 'Arial, Helvetica, sans-serif');
    const baseSize = String(opts.baseSize || '10pt');
    const baseLineHeight = String(opts.baseLineHeight || '1.45');
    let html = String(input || '');
    if (!html) return '';

    const strip = (frag) => String(frag || '')
      .replace(/font-size\s*:\s*[^;"']+;?/gi, '')
      .replace(/\sstyle=\"\s*\"/gi, '');
    const wrap = (frag) => {
      const t = String(frag || '').trim();
      if (!t) return '';
      return `<div style="font-family:${baseFamily};font-size:${baseSize};line-height:${baseLineHeight};color:#3c3c3b;">${t}</div>`;
    };

    const start = html.indexOf(SIG_START);
    const end = start >= 0 ? html.indexOf(SIG_END, start + SIG_START.length) : -1;
    if (start >= 0 && end > start) {
      const before = html.slice(0, start);
      const sig = html.slice(start, end + SIG_END.length);
      const after = html.slice(end + SIG_END.length);
      return `${wrap(strip(before))}${sig}${wrap(strip(after))}`;
    }
    return wrap(strip(html));
  } catch (_err) {
    return String(input || '');
  }
}
const DEFAULT_FIELD_VALUE_SOURCES = {
  client_name: 'jobsheet.client_name',
  client_email: 'jobsheet.client_email',
  client_phone: 'jobsheet.client_phone',
  client_address1: 'jobsheet.client_address1',
  client_address2: 'jobsheet.client_address2',
  client_address3: 'jobsheet.client_address3',
  client_town: 'jobsheet.client_town',
  client_postcode: 'jobsheet.client_postcode',
  event_type: 'jobsheet.event_type',
  event_date: 'jobsheet.event_date',
  event_start: 'jobsheet.event_start',
  event_end: 'jobsheet.event_end',
  venue_name: 'jobsheet.venue_name',
  venue_address1: 'jobsheet.venue_address1',
  venue_address2: 'jobsheet.venue_address2',
  venue_address3: 'jobsheet.venue_address3',
  venue_town: 'jobsheet.venue_town',
  venue_postcode: 'jobsheet.venue_postcode',
  caterer_name: 'jobsheet.caterer_name',
  // AhMen singer fee should come from the jobsheet-sourced value
  ahmen_fee: 'jobsheet.ahmen_fee',
  total_amount: 'context.totalAmount',
  extra_fees: 'context.extraFees',
  production_fees: 'context.productionFees',
  deposit_amount: 'context.depositAmount',
  balance_amount: 'context.balanceAmount',
  balance_due_date: 'context.balanceDate',
  balance_reminder_date: 'context.balanceRemind',
  service_types: 'jobsheet.service_types',
  specialist_singers: 'jobsheet.specialist_singers'
};

function resolvePath(targetPath) {
  if (!targetPath || typeof targetPath !== 'string') {
    throw new Error('Template path is required');
  }
  const normalized = targetPath.replace(/^~\//, `${process.env.HOME || ''}/`);
  const resolved = path.resolve(normalized);
  return resolved;
}

function normalizeRecipientList(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.flatMap(item => normalizeRecipientList(item));
  }
  return String(value)
    .split(/[,;]+/)
    .map(part => part.trim())
    .filter(Boolean);
}

function parseStoredRecipients(value) {
  if (!value) return [];
  return String(value)
    .split(/[,;]+/)
    .map(part => part.trim())
    .filter(Boolean);
}

function parseStoredAttachments(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed.filter(Boolean).map(String);
    return [];
  } catch (_) {
    return [];
  }
}

function broadcastJobsheetChange(payload) {
  if (!IS_MAIN_PROCESS) return;
  try {
    if (!ElectronBrowserWindow) {
      ({ BrowserWindow: ElectronBrowserWindow } = require('electron'));
    }
    const message = payload || {};
    (ElectronBrowserWindow.getAllWindows() || []).forEach(win => {
      if (!win || win.isDestroyed()) return;
      try {
        win.webContents.send('jobsheet-change', message);
      } catch (err) {
        console.warn('broadcastJobsheetChange failed for a window', err);
      }
    });
  } catch (err) {
    console.warn('Unable to broadcast jobsheet change', err);
  }
}

function ensureScheduledMailWorker() {
  if (!IS_MAIN_PROCESS) return;
  if (scheduledMailWorkerStarted) return;
  scheduledMailWorkerStarted = true;

  const tick = async () => {
    if (scheduledMailWorkerExecuting) return;
    scheduledMailWorkerExecuting = true;
    try {
      const due = await db.listDueScheduledEmails({ limit: SCHEDULED_EMAIL_BATCH_SIZE });
      for (const item of due) {
        // eslint-disable-next-line no-await-in-loop
        await processScheduledEmail(item);
      }
    } catch (err) {
      console.error('Scheduled email worker error', err);
    } finally {
      scheduledMailWorkerExecuting = false;
    }
  };

  tick();
  setInterval(tick, SCHEDULED_EMAIL_POLL_INTERVAL_MS);
}

async function processScheduledEmail(entry) {
  const attachments = parseStoredAttachments(entry.attachments);
  const ccList = parseStoredRecipients(entry.cc_address);
  const bccList = parseStoredRecipients(entry.bcc_address);
  const payload = {
    to: entry.to_address,
    cc: ccList,
    bcc: bccList,
    subject: entry.subject,
    body: entry.body,
    attachments,
    is_html: entry.is_html === 1 || entry.is_html === true,
    business_id: entry.business_id,
    jobsheet_id: entry.jobsheet_id,
    skipLog: true
  };

  try {
    await sendMailViaGraph(payload);
    await db.markScheduledEmailSent({ id: entry.id, sent_at: new Date() });
    if (entry.email_log_id) {
      await db.updateEmailLogStatus({ id: entry.email_log_id, status: 'sent', sent_at: new Date() });
    }
    broadcastJobsheetChange({
      type: 'email-log-updated',
      businessId: entry.business_id != null ? Number(entry.business_id) : null,
      jobsheetId: entry.jobsheet_id != null ? Number(entry.jobsheet_id) : null
    });
  } catch (err) {
    console.error('Scheduled email send failed', err);
    const attempts = Number(entry.attempt_count) || 0;
    const delayMinutes = Math.min(60, Math.max(5, (attempts + 1) * 5));
    await db.markScheduledEmailFailed({ id: entry.id, error: err.message, retryInMinutes: delayMinutes });
    if (entry.email_log_id) {
      await db.updateEmailLogStatus({ id: entry.email_log_id, status: 'scheduled_error' });
    }
    broadcastJobsheetChange({
      type: 'email-log-updated',
      businessId: entry.business_id != null ? Number(entry.business_id) : null,
      jobsheetId: entry.jobsheet_id != null ? Number(entry.jobsheet_id) : null
    });
  }
}

const SPLIT_WORKBOOKS_DIR = process.env.SPLIT_WORKBOOKS_DIR || '/Users/motticohen/Dropbox/My Invoicing App/AhMen/TEMPLATES';

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Graceful Excel shutdown scheduler so batch exports can keep Excel open
let excelQuitTimer = null;
let excelShouldQuitAtEnd = null; // set per batch based on whether Excel was running at batch start

async function isExcelRunning() {
  return await new Promise(resolve => {
    try {
      execFile('osascript', ['-e', 'tell application "System Events" to (exists process "Microsoft Excel")'], (error, stdout) => {
        if (error) return resolve(false);
        const text = (stdout || '').toString().trim().toLowerCase();
        resolve(text === 'true' || text === 'yes');
      });
    } catch (_err) {
      resolve(false);
    }
  });
}

function scheduleExcelQuit(delayMs = 10000) {
  try { if (excelQuitTimer) clearTimeout(excelQuitTimer); } catch (_) {}
  if (!excelShouldQuitAtEnd) {
    // Do not schedule quit if Excel was already running at batch start
    return;
  }
  excelQuitTimer = setTimeout(() => {
    try {
      execFile('osascript', ['-e', 'tell application "Microsoft Excel" to quit saving no'], () => {});
    } catch (_) { /* ignore */ }
    // Reset batch state after quit attempt
    try { if (excelQuitTimer) clearTimeout(excelQuitTimer); } catch (_) {}
    excelQuitTimer = null;
    excelShouldQuitAtEnd = null;
  }, Math.max(0, Number(delayMs) || 0));
}

async function pathExists(targetPath) {
  if (!targetPath) return false;
  try {
    await fs.promises.access(targetPath);
    return true;
  } catch (err) {
    return false;
  }
}

function isSubPath(parentPath, childPath) {
  const parentResolved = path.resolve(parentPath);
  const childResolved = path.resolve(childPath);
  const relative = path.relative(parentResolved, childResolved);
  return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative));
}

async function ensureFileAccessible(resolvedPath) {
  await fs.promises.access(resolvedPath, fs.constants.R_OK);
}

async function waitForFile(targetPath, timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await pathExists(targetPath)) {
      return true;
    }
    await sleep(300);
  }
  return false;
}

async function normalizeTemplate(args = {}) {
  const rawPath = args.templatePath || args.path;
  const resolvedPath = resolvePath(rawPath);
  await ensureFileAccessible(resolvedPath);

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(resolvedPath);
  } catch (err) {
    throw new Error(`Unable to read workbook: ${err.message || err}`);
  }

  return {
    ok: true,
    templatePath: resolvedPath
  };
}

function sanitizeFilenameSegment(value) {
  if (!value) return '';
  return value
    .toString()
    .replace(INVALID_FILENAME_CHARS, ' ')
    .replace(/[\s]+/g, ' ')
    .trim();
}

function formatDateISO(dateInput) {
  if (!dateInput) return '';
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return sanitizeFilenameSegment(dateInput);
  return date.toISOString().slice(0, 10);
}

function formatDateHuman(dateInput) {
  if (!dateInput) return '';
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return dateInput;
  return new Intl.DateTimeFormat('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  }).format(date);
}

async function ensureDirectoryExists(targetPath) {
  if (!targetPath) return;
  await fs.promises.mkdir(targetPath, { recursive: true });
}

function extractPrefixFromInvoiceCell(value) {
  if (value == null) return 'INV-';
  const text = String(value).trim();
  if (!text) return 'INV-';
  // Take all leading non-digit characters (preserve hyphens and spaces)
  const match = text.match(/^([^0-9]+)/);
  const prefix = match ? match[1].trim() : '';
  return prefix || 'INV-';
}

async function readInvoicePrefixFromWorkbook(workbookPath) {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(workbookPath);
    // Prefer an invoice sheet if present
    let ws = wb.worksheets.find(s => /invoice/i.test(String(s.name || '')));
    if (!ws) ws = wb.worksheets && wb.worksheets.length ? wb.worksheets[0] : null;
    if (!ws) return 'INV-';
    const cell = ws.getCell('E9');
    const val = cell ? cell.value : null;
    return extractPrefixFromInvoiceCell(val);
  } catch (_err) {
    return 'INV-';
  }
}

// Extract jobsheet data from a folder containing legacy files
function parseFolderNameForJob(folderName) {
  if (!folderName) return { client_name: '', event_date: '' };
  const name = String(folderName).trim();
  // Pattern: YYYY-MM-DD - Client Name - Human date (last part optional)
  const re = /^(\d{4}-\d{2}-\d{2})\s*-\s*([^\-]+?)(?:\s*-\s*.*)?$/;
  const m = name.match(re);
  if (!m) return { client_name: '', event_date: '' };
  const event_date = m[1] || '';
  const client_name = (m[2] || '').trim();
  return { client_name, event_date };
}

async function tryReadClientDataSheet(workbookPath) {
  if (!workbookPath) return {};
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(workbookPath);
    const ws = wb.getWorksheet('Client Data') || wb.worksheets.find(s => /client\s*data/i.test(String(s.name || '')));
    if (!ws) return {};
    const val = (addr) => {
      try { const cell = ws.getCell(addr); return cell ? cell.value : undefined; } catch (_) { return undefined; }
    };
    const num = (addr) => {
      const v = val(addr);
      const n = typeof v === 'object' && v && v.result != null ? Number(v.result) : Number(v);
      return Number.isFinite(n) ? n.toFixed(2) : '';
    };
    const date = (addr) => {
      const v = val(addr);
      if (!v) return '';
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      const asNum = Number(v);
      if (Number.isFinite(asNum) && asNum > 0) {
        // Excel serial date (rough; ExcelJS often gives JS Date already)
        const js = new Date(Math.round((asNum - 25569) * 86400 * 1000));
        if (!Number.isNaN(js.valueOf())) return js.toISOString().slice(0, 10);
      }
      const js = new Date(v);
      return Number.isNaN(js.valueOf()) ? '' : js.toISOString().slice(0, 10);
    };
    const str = (addr) => {
      const v = val(addr);
      if (v == null) return '';
      if (typeof v === 'object' && v && v.text) return String(v.text);
      return String(v);
    };
    return {
      // Client
      client_name: str('B3'),
      client_email: str('B4'),
      client_phone: str('B5'),
      client_address1: str('B6'),
      client_address2: str('B7'),
      client_address3: str('B8'),
      client_town: str('B9'),
      client_postcode: str('B10'),
      // Event
      event_type: str('B13'),
      event_date: date('B14'),
      event_start: str('B15'),
      event_end: str('B16'),
      // Venue
      venue_name: str('B19'),
      venue_address1: str('B20'),
      venue_address2: str('B21'),
      venue_address3: str('B22'),
      venue_town: str('B23'),
      venue_postcode: str('B24'),
      // Financials (advisory; editor derives deposit/balance later)
      ahmen_fee: num('B27'),
      extra_fees: num('B28'),
      production_fees: num('B29'),
      total_amount: num('B30'),
      deposit_amount: num('B31'),
      balance_amount: num('B32'),
      balance_due_date: date('B33'),
      balance_reminder_date: date('B34'),
      // Services / notes
      service_types: str('B37'),
      specialist_singers: str('B38'),
      caterer_name: str('B40')
    };
  } catch (_err) {
    return {};
  }
}

async function extractJobsheetDataFromFolder(options = {}) {
  const folderPath = options.folderPath || options.path;
  if (!folderPath) return { ok: false, message: 'folderPath is required' };
  const resolved = path.resolve(folderPath);
  let stat = null;
  try { stat = await fs.promises.stat(resolved); } catch (err) { return { ok: false, message: 'Folder not found' }; }
  if (!stat.isDirectory()) return { ok: false, message: 'Path is not a folder' };

  const base = path.basename(resolved);
  const fromName = parseFolderNameForJob(base);

  let entries = [];
  try { entries = await fs.promises.readdir(resolved); } catch (_) { entries = []; }
  const excelCandidates = entries.filter(n => /\.xlsx$/i.test(n)).map(n => path.join(resolved, n));
  let workbookPath = null;
  let workbookFields = {};
  for (const p of excelCandidates) {
    const fields = await tryReadClientDataSheet(p); // eslint-disable-line no-await-in-loop
    const hasAny = Object.values(fields).some(v => v != null && v !== '');
    if (hasAny) {
      workbookPath = p;
      workbookFields = fields;
      break;
    }
  }

  // collect invoice PDFs with (INV-###)
  const pdfs = entries.filter(n => /\.pdf$/i.test(n)).map(n => path.join(resolved, n));
  const invoices = pdfs.filter(fp => /\(\s*INV[-\s]?\d+\s*\)\.pdf$/i.test(path.basename(fp)));

  // Build suggested values excluding financial fields (manual entry in jobsheet)
  const FINANCIAL_KEYS = new Set([
    'ahmen_fee','extra_fees','production_fees','total_amount','deposit_amount','balance_amount','balance_due_date','balance_reminder_date'
  ]);
  const suggested = {};
  Object.entries(workbookFields || {}).forEach(([k, v]) => {
    if (!FINANCIAL_KEYS.has(k) && v !== undefined && v !== '') suggested[k] = v;
  });
  if (fromName.client_name && !suggested.client_name) suggested.client_name = fromName.client_name;
  if (fromName.event_date && !suggested.event_date) suggested.event_date = fromName.event_date;

  return {
    ok: true,
    folder: resolved,
    workbook_path: workbookPath,
    invoices,
    suggested
  };
}

async function createStampedWorkbookCopy(sourcePath, stampText) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(sourcePath);
  const targets = wb.worksheets.filter(s => /invoice/i.test(String(s.name || '')));
  const list = targets.length ? targets : wb.worksheets.slice(0, 1);
  list.forEach(ws => {
    try {
      ws.getCell('E9').value = stampText;
    } catch (_err) {}
  });
  const dir = path.dirname(sourcePath);
  const tmpPath = path.join(dir, `.inv_tmp_${Date.now()}_${Math.random().toString(36).slice(2)}.xlsx`);
  await wb.xlsx.writeFile(tmpPath);
  return tmpPath;
}

function buildContext(payload, business = {}) {
  const jobsheet = { ...(payload.jobsheet_snapshot || {}) };
  const client = { ...payload.client_override };
  const event = { ...payload.event_override };
  const pricing = { ...(payload.pricing_snapshot || {}) };

  const derived = {
    totalAmount: payload.total_amount ?? payload.balance_amount ?? payload.balance_due ?? null,
    extraFees: payload.extra_fees ?? null,
    productionFees: payload.production_fees ?? null,
    depositAmount: payload.deposit_amount ?? null,
    balanceAmount: payload.balance_amount ?? payload.balance_due ?? null,
    balanceDate: payload.balance_due_date ?? null,
    balanceRemind: payload.balance_reminder_date ?? null
  };

  return {
    jobsheet,
    client,
    event,
    pricing,
    business,
    document: payload,
    context: derived
  };
}

function resolvePathValue(root, pathExpression) {
  if (!pathExpression || !root) return undefined;
  const parts = pathExpression.split('.').map(part => part.trim()).filter(Boolean);
  let current = root;
  for (const part of parts) {
    if (current == null) return undefined;
    current = current[part];
  }
  return current;
}

function resolveFieldValue(fieldKey, valueSources, context, fallbackPath) {
  const source = valueSources[fieldKey] || null;
  if (source) {
    if (source.source_type === 'literal') {
      return source.literal_value;
    }
    if (source.source_type === 'contextPath' && source.source_path) {
      const value = resolvePathValue(context, source.source_path);
      if (value !== undefined) {
        return value;
      }
    }
  }

  const defaultPath = fallbackPath || DEFAULT_FIELD_VALUE_SOURCES[fieldKey];
  if (defaultPath) {
    return resolvePathValue(context, defaultPath);
  }

  return undefined;
}

function toExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Number.isNaN(value.valueOf()) ? null : value;
  }
  const dt = new Date(value);
  return Number.isNaN(dt.valueOf()) ? null : dt;
}

function coerceCellValue(rawValue, binding) {
  if (rawValue === undefined || rawValue === null || rawValue === '') {
    return '';
  }

  if (typeof rawValue === 'number' && !Number.isFinite(rawValue)) {
    return null;
  }

  if (typeof rawValue === 'string') {
    const normalized = rawValue.trim().toLowerCase();
    if (normalized === 'nan' || normalized === 'infinity' || normalized === '-infinity') {
      return '';
    }
  }

  const { data_type: dataType = 'string', format } = binding || {};

  // Prefer true Excel date (Date object) when data_type requests a date
  if (dataType === 'date') {
    const d = toExcelDate(rawValue);
    return d || '';
  }

  // Legacy/support: allow explicit request for a human string
  if (format === 'date_human') {
    return formatDateHuman(rawValue);
  }

  if (dataType === 'number') {
    let numeric;
    if (typeof rawValue === 'string') {
      const cleaned = rawValue.replace(/[^0-9.\-]+/g, '');
      numeric = Number(cleaned);
    } else {
      numeric = Number(rawValue);
    }
    return Number.isFinite(numeric) ? numeric : null;
  }

  if (typeof rawValue === 'number' && Number.isNaN(rawValue)) {
    return '';
  }

  if (dataType === 'string') {
    return rawValue.toString();
  }

  return rawValue;
}

function applyNumberFormat(cell, binding) {
  if (!cell || !binding) return;
  const fmt = (binding.format || '').toLowerCase();
  const dataType = (binding.data_type || '').toLowerCase();
  let numFmt = null;
  if (dataType === 'date') {
    // Enforce long UK-style date: 15 October 2025
    numFmt = 'dd mmmm yyyy';
  } else if (dataType === 'number') {
    if (fmt === 'percentage' || fmt === 'percent') numFmt = '0.00%';
    else if (fmt === 'integer' || fmt === 'whole') numFmt = '0';
    else if (fmt === 'decimal_2' || fmt === 'decimal2' || fmt === 'number_2dp') numFmt = '#,##0.00';
    else if (fmt === 'currency' || fmt === 'money' || fmt === 'gbp' || fmt === '£') numFmt = '£#,##0.00';
  }
  if (!numFmt) return;
  try {
    // Always enforce date format. For numbers, override when format is currency
    // or when the existing format is blank, General, or Text ('@').
    if (dataType === 'date') {
      cell.numFmt = numFmt;
    } else {
      const existing = (cell.numFmt || '').toString().toLowerCase();
      const shouldOverride = !existing || existing === 'general' || existing === '@' || fmt === 'currency' || fmt === 'money' || fmt === 'gbp' || fmt === '£';
      if (shouldOverride) {
        cell.numFmt = numFmt;
      }
    }
  } catch (_err) {}
}

async function fillWorkbook(workbook, bindings, valueSources, context) {
  if (!Array.isArray(bindings) || !bindings.length) return;

  const fallbackPaths = DEFAULT_FIELD_VALUE_SOURCES;

  bindings.forEach(binding => {
    if (!binding || !binding.sheet || !binding.cell || !binding.field_key) return;

    const worksheet = workbook.getWorksheet(binding.sheet);
    if (!worksheet) return;

    const cell = worksheet.getCell(binding.cell);
    if (!cell) return;

    const value = resolveFieldValue(binding.field_key, valueSources, context, fallbackPaths[binding.field_key]);
    const coerced = coerceCellValue(value, binding);
    cell.value = coerced;
    applyNumberFormat(cell, binding);
  });
}

function collectPlaceholderKeys(workbook) {
  const keys = new Set();
  workbook.eachSheet(worksheet => {
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        const value = cell?.value;
        if (typeof value === 'string') {
          let match;
          PLACEHOLDER_PATTERN.lastIndex = 0;
          while ((match = PLACEHOLDER_PATTERN.exec(value)) !== null) {
            if (match[1]) keys.add(match[1]);
          }
        } else if (value && typeof value === 'object' && Array.isArray(value.richText)) {
          value.richText.forEach(fragment => {
            if (!fragment?.text) return;
            let match;
            PLACEHOLDER_PATTERN.lastIndex = 0;
            while ((match = PLACEHOLDER_PATTERN.exec(fragment.text)) !== null) {
              if (match[1]) keys.add(match[1]);
            }
          });
        }
      });
    });
  });
  return keys;
}

function replaceWorkbookPlaceholders(workbook, valueSources, context, placeholderMap = new Map()) {
  const fallbackPaths = DEFAULT_FIELD_VALUE_SOURCES;

  // Keys that should be rendered as currency (GBP) when used as placeholders
  const CURRENCY_KEYS = new Set([
    'ahmen_fee',
    'total_amount',
    'extra_fees',
    'production_fees',
    'deposit_amount',
    'balance_amount'
  ]);

  // Keys that represent dates
  const DATE_KEYS = new Set([
    'event_date',
    'balance_due_date',
    'balance_reminder_date',
    'document_date'
  ]);

  const formatCurrency = (val) => {
    const num = Number(val);
    if (!Number.isFinite(num)) return '';
    try {
      return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(num);
    } catch (_err) {
      // Fallback if Intl fails for any reason
      return `£${num.toFixed(2)}`;
    }
  };

  const toNumeric = (val) => {
    if (val === null || val === undefined) return null;
    if (typeof val === 'number') return Number.isFinite(val) ? val : null;
    const cleaned = String(val).replace(/[^0-9.\-]+/g, '');
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
  };

  const resolvePlaceholder = (rawKey) => {
    const lookup = String(rawKey || '').toLowerCase();
    const slug = normalizeTokenKey(lookup);
    const fieldKey = placeholderMap.get(lookup) || placeholderMap.get(slug) || rawKey;
    return resolveFieldValue(fieldKey, valueSources, context, fallbackPaths[fieldKey]);
  };

  const renderValueForPlaceholder = (fieldKey, value) => {
    const keyLower = String(fieldKey || '').toLowerCase();
    if (value === undefined || value === null) return '';
    if (value instanceof Date) return formatDateHuman(value) || '';
    if (DATE_KEYS.has(keyLower)) return formatDateHuman(value);
    if (CURRENCY_KEYS.has(keyLower)) return formatCurrency(value);
    if (typeof value === 'number' && Number.isFinite(value)) return value.toString();
    return value != null ? value.toString() : '';
  };

  const tokenKeys = Array.from(placeholderMap.keys());

  workbook.eachSheet(worksheet => {
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        const current = cell?.value;
        if (typeof current === 'string') {
          // If the entire cell is a single placeholder like {{AHMEN_FEE}}, write a typed value
          const singleToken = current.trim().match(/^{{\s*([a-zA-Z0-9_.-]+)\s*}}$/);
          if (singleToken) {
            const rawKey = singleToken[1];
            const mappedKey = placeholderMap.get(String(rawKey).toLowerCase()) || rawKey;
            const resolved = resolvePlaceholder(rawKey);
            if (resolved === undefined || resolved === null) {
              cell.value = '';
            } else if (DATE_KEYS.has(String(mappedKey).toLowerCase())) {
              const dt = toExcelDate(resolved);
              if (dt) {
                cell.value = dt;
                applyNumberFormat(cell, { data_type: 'date' });
              } else {
                cell.value = formatDateHuman(resolved) || '';
              }
            } else if (typeof resolved === 'number' && Number.isFinite(resolved)) {
              cell.value = resolved;
              if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
                applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
              }
            } else {
              // Attempt to coerce numeric strings for currency keys
              if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
                const n = toNumeric(resolved);
                if (n !== null) {
                  cell.value = n;
                  applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
                } else {
                  cell.value = renderValueForPlaceholder(mappedKey, resolved);
                }
              } else {
                cell.value = renderValueForPlaceholder(mappedKey, resolved);
              }
            }
            return; // handled this cell
          }
          PLACEHOLDER_PATTERN.lastIndex = 0;
          const updated = current.replace(PLACEHOLDER_PATTERN, (match, key) => {
            if (!key) return '';
            const mappedKey = placeholderMap.get(String(key).toLowerCase()) || placeholderMap.get(normalizeTokenKey(key)) || key;
            const resolved = resolvePlaceholder(key);
            return renderValueForPlaceholder(mappedKey, resolved);
          });
          if (updated !== current) {
            cell.value = updated;
          } else {
            // Fallback: exact token match without braces
            const trimmedLower = current.trim().toLowerCase();
            if (tokenKeys.includes(trimmedLower)) {
              const fieldKey = placeholderMap.get(trimmedLower) || placeholderMap.get(normalizeTokenKey(trimmedLower)) || trimmedLower;
              const resolved = resolvePlaceholder(fieldKey);
              if (resolved === undefined || resolved === null) {
                cell.value = '';
              } else if (resolved instanceof Date) {
                if (DATE_KEYS.has(String(fieldKey).toLowerCase())) {
                  cell.value = toExcelDate(resolved) || '';
                  applyNumberFormat(cell, { data_type: 'date' });
                } else {
                  cell.value = formatDateHuman(resolved) || '';
                }
              } else if (typeof resolved === 'number' && Number.isFinite(resolved)) {
                if (CURRENCY_KEYS.has(String(fieldKey).toLowerCase())) {
                  // Keep numeric value for Excel but enforce currency formatting
                  cell.value = resolved;
                  applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
                } else {
                  cell.value = resolved;
                }
              } else {
                if (DATE_KEYS.has(String(fieldKey).toLowerCase())) {
                  const dt = toExcelDate(resolved);
                  if (dt) {
                    cell.value = dt;
                    applyNumberFormat(cell, { data_type: 'date' });
                  } else {
                    cell.value = formatDateHuman(resolved);
                  }
                } else {
                  cell.value = resolved.toString();
                }
              }
            }
          }
        } else if (current && typeof current === 'object' && Array.isArray(current.richText)) {
          // If the entire richText content is a single placeholder token, write a typed value
          const fullText = current.richText.map(f => (f && typeof f.text === 'string' ? f.text : '')).join('');
          const singleToken = fullText.trim().match(/^{{\s*([a-zA-Z0-9_.-]+)\s*}}$/);
          if (singleToken) {
            const rawKey = singleToken[1];
            const mappedKey = placeholderMap.get(String(rawKey).toLowerCase()) || rawKey;
            const resolved = resolvePlaceholder(rawKey);
            if (resolved === undefined || resolved === null) {
              cell.value = '';
            } else if (DATE_KEYS.has(String(mappedKey).toLowerCase())) {
              const dt = toExcelDate(resolved);
              if (dt) {
                cell.value = dt;
                applyNumberFormat(cell, { data_type: 'date' });
              } else {
                cell.value = formatDateHuman(resolved) || '';
              }
            } else if (typeof resolved === 'number' && Number.isFinite(resolved)) {
              cell.value = resolved;
              if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
                applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
              }
            } else {
              if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
                const n = toNumeric(resolved);
                if (n !== null) {
                  cell.value = n;
                  applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
                } else {
                  cell.value = renderValueForPlaceholder(mappedKey, resolved);
                }
              } else {
                cell.value = renderValueForPlaceholder(mappedKey, resolved);
              }
            }
            return; // handled this cell
          }
          // Also handle bare token without braces for the whole richText
          const bare = fullText.trim().toLowerCase();
          if (tokenKeys.includes(bare)) {
            const mappedKey = placeholderMap.get(bare) || bare;
            const resolved = resolvePlaceholder(mappedKey);
            if (resolved === undefined || resolved === null) {
              cell.value = '';
            } else if (DATE_KEYS.has(String(mappedKey).toLowerCase())) {
              const dt = toExcelDate(resolved);
              if (dt) {
                cell.value = dt;
                applyNumberFormat(cell, { data_type: 'date' });
              } else {
                cell.value = formatDateHuman(resolved) || '';
              }
            } else if (typeof resolved === 'number' && Number.isFinite(resolved)) {
              cell.value = resolved;
              if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
                applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
              }
            } else if (CURRENCY_KEYS.has(String(mappedKey).toLowerCase())) {
              const n = toNumeric(resolved);
              if (n !== null) {
                cell.value = n;
                applyNumberFormat(cell, { data_type: 'number', format: 'currency' });
              } else {
                cell.value = renderValueForPlaceholder(mappedKey, resolved);
              }
            } else {
              cell.value = renderValueForPlaceholder(mappedKey, resolved);
            }
            return; // handled
          }
          let changed = false;
          const richText = current.richText.map(fragment => {
            if (!fragment?.text) return fragment;
            const original = fragment.text;
            PLACEHOLDER_PATTERN.lastIndex = 0;
            const updated = original.replace(PLACEHOLDER_PATTERN, (match, key) => {
              if (!key) return '';
              const mappedKey = placeholderMap.get(String(key).toLowerCase()) || placeholderMap.get(normalizeTokenKey(key)) || key;
              const resolved = resolvePlaceholder(key);
              return renderValueForPlaceholder(mappedKey, resolved);
            });
            if (updated !== original) {
              changed = true;
              return { ...fragment, text: updated };
            }
            // Fallback: exact token match without braces
            const trimmedLower = original.trim().toLowerCase();
            if (tokenKeys.includes(trimmedLower)) {
              const fieldKey = placeholderMap.get(trimmedLower) || placeholderMap.get(normalizeTokenKey(trimmedLower)) || trimmedLower;
              const resolved = resolvePlaceholder(fieldKey);
              changed = true;
              const textVal = renderValueForPlaceholder(fieldKey, resolved);
              return { ...fragment, text: textVal };
            }
            return fragment;
          });
          if (changed) {
            cell.value = { ...current, richText };
          }
        }
      });
    });
  });
}

function sanitizeWorkbookValues(workbook) {
  workbook.eachSheet(worksheet => {
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        const { value } = cell;
        if (value == null) return;

        if (value instanceof Date) {
          if (Number.isNaN(value.valueOf())) {
            cell.value = null;
          }
          return;
        }

        if (typeof value === 'number' && !Number.isFinite(value)) {
          cell.value = null;
          return;
        }

        if (typeof value === 'string') {
          const normalized = value.trim().toLowerCase();
          if (normalized === 'nan' || normalized === 'infinity' || normalized === '-infinity') {
            cell.value = '';
          }
          return;
        }

        if (typeof value === 'object' && value !== null && ('formula' in value || 'sharedFormula' in value)) {
          const result = value.result;
          let sanitizedResult = result;

          if (result instanceof Date && Number.isNaN(result.valueOf())) {
            sanitizedResult = null;
          } else if (typeof result === 'number' && !Number.isFinite(result)) {
            sanitizedResult = null;
          } else if (typeof result === 'string') {
            const normalized = result.trim().toLowerCase();
            if (normalized === 'nan' || normalized === 'infinity' || normalized === '-infinity') {
              sanitizedResult = null;
            }
          }

          if (sanitizedResult !== result) {
            cell.value = {
              ...value,
              result: sanitizedResult
            };
          }
        }
      });
    });
  });
}

async function saveWorkbookAsPdf(sourcePath, targetPath, _options = {}) {
  await fs.promises.rm(targetPath, { force: true }).catch(() => {});

  // If no batch timer is active, capture whether Excel is currently running
  if (excelQuitTimer == null && excelShouldQuitAtEnd == null) {
    try {
      const running = await isExcelRunning();
      excelShouldQuitAtEnd = !running; // quit later only if not already running
    } catch (_err) {
      excelShouldQuitAtEnd = false;
    }
  }

  // Primary: simple AppleScript using Excel's PDF file format
  const { activeSheetOnly } = _options || {};
  const stampCell = (_options && _options.stampCell) ? String(_options.stampCell) : '';
  const stampText = (_options && _options.stampText) ? String(_options.stampText) : '';
  const stampSheetName = (_options && _options.stampSheetName) ? String(_options.stampSheetName) : '';
  const stampVariant = (_options && _options.stampVariant) ? String(_options.stampVariant) : '';

  const osaArgs = [
    '-e', 'on run argv',
    '-e', 'if (count of argv) < 2 then error "Missing arguments"',
    '-e', 'set workbookPosixPath to item 1 of argv',
    '-e', 'set targetPdfPosix to item 2 of argv',
    // Optional args for stamping
    '-e', 'set stampCell to missing value',
    '-e', 'set stampText to missing value',
    '-e', 'set stampSheetName to missing value',
    '-e', 'set stampVariant to missing value',
    '-e', 'if (count of argv) ≥ 3 then set stampCell to item 3 of argv',
    '-e', 'if (count of argv) ≥ 4 then set stampText to item 4 of argv',
    '-e', 'if (count of argv) ≥ 5 then set stampSheetName to item 5 of argv',
    '-e', 'if (count of argv) ≥ 6 then set stampVariant to item 6 of argv',
    '-e', 'set workbookHfs to (POSIX file workbookPosixPath) as text',
    '-e', 'set targetPdfHfs to (POSIX file targetPdfPosix) as text',
    '-e', 'set pdfAlias to POSIX file targetPdfPosix',
    '-e', 'tell application "Microsoft Excel"',
    '-e', 'launch',
    // Keep Excel in background: no activate, hide UI and alerts
    '-e', 'try',
    '-e', 'set visible to false',
    '-e', 'set display alerts to false',
    '-e', 'end try',
    '-e', 'set wb to missing value',
    '-e', 'try',
    '-e', 'set wb to open workbook workbook file name workbookHfs',
    '-e', 'end try',
    '-e', 'repeat with i from 1 to 50',
    '-e', 'if wb is not missing value then exit repeat',
    '-e', 'delay 0.1',
    '-e', 'try',
    '-e', 'set wb to active workbook',
    '-e', 'end try',
    '-e', 'end repeat',
    '-e', 'if wb is missing value then error "Unable to open workbook"',
    // If stamping info was provided, update the cell value on the target or active sheet before export
    '-e', 'if stampCell is not missing value and stampText is not missing value then',
    '-e', 'try',
    '-e', 'set theSheet to active sheet of wb',
    '-e', 'if stampSheetName is not missing value then',
    '-e', 'try',
    '-e', 'set theSheet to worksheet stampSheetName of wb',
    '-e', 'end try',
    '-e', 'end if',
    '-e', 'if stampVariant is not missing value then',
    '-e', 'if theSheet is missing value then set theSheet to active sheet of wb',
    '-e', 'try',
    '-e', 'set found to false',
    '-e', 'repeat with ws in (worksheets of wb)',
    '-e', 'set nm to (name of ws) as text',
    '-e', 'ignoring case',
    '-e', 'if (nm contains "invoice") and (nm contains stampVariant) then',
    '-e', 'set theSheet to ws',
    '-e', 'set found to true',
    '-e', 'exit repeat',
    '-e', 'end if',
    '-e', 'end ignoring',
    '-e', 'end repeat',
    '-e', 'end try',
    '-e', 'end if',
    '-e', 'set value of range stampCell of theSheet to stampText',
    '-e', 'end try',
    '-e', 'end if',
    '-e', 'set wbName to name of wb',
    '-e', 'delay 0.2',
    '-e', 'try',
    '-e', 'save workbook as wb filename targetPdfHfs file format PDF file format',
    '-e', 'on error errMsg number errNum',
    '-e', 'try',
    '-e', 'close workbook wb saving no',
    '-e', 'end try',
    '-e', 'error errMsg number errNum',
    '-e', 'end try',
    // Ensure workbook closes even if Excel is busy
    '-e', 'try',
    '-e', 'close workbook wb saving no',
    '-e', 'end try',
    '-e', 'try',
    '-e', 'tell wb to close saving no',
    '-e', 'end try',
    // Fallback: close any workbook with the same name (AppleScript, not VB)
    '-e', 'try',
    '-e', 'repeat with bk in (workbooks whose name is wbName)',
    '-e', 'close bk saving no',
    '-e', 'end repeat',
    '-e', 'end try',
    // Retry-close loop in case Excel is momentarily busy after export
    '-e', 'repeat with i from 1 to 20',
    '-e', 'try',
    '-e', 'close workbook wb saving no',
    '-e', 'exit repeat',
    '-e', 'on error errMsg number errNum',
    '-e', 'delay 0.2',
    '-e', 'end try',
    '-e', 'end repeat',
    '-e', 'delay 0.1',
    '-e', 'end tell',
    // Ensure Excel is not frontmost as a fallback
    '-e', 'try',
    '-e', 'tell application "System Events" to set frontmost of process "Microsoft Excel" to false',
    '-e', 'end try',
    '-e', 'end run',
    sourcePath,
    targetPath
  ];

  // Pass optional stamping args if provided
  if (stampCell && stampText) {
    osaArgs.push(stampCell);
    osaArgs.push(stampText);
    if (stampSheetName) osaArgs.push(stampSheetName); else osaArgs.push('');
    if (stampVariant) osaArgs.push(stampVariant);
  }

  await new Promise((resolve, reject) => {
    execFile('osascript', osaArgs, { timeout: 120000 }, (error, stdout, stderr) => {
      if (error) {
        const raw = (stderr || stdout || '').toString();
        const message = `${error.message || 'Unable to export workbook to PDF'}${raw ? `\n${raw.trim()}` : ''}`.trim();
        reject(new Error(message));
        return;
      }
      resolve();
    });
  });

  const created = await waitForFile(targetPath, 60000);
  if (!created) {
    throw new Error(`PDF not found after export: ${targetPath}`);
  }

  // Schedule Excel to quit shortly after export; repeated exports will reset the timer
  scheduleExcelQuit(10000);
}

function parseWorkbookName(filePath) {
  const baseName = path.basename(filePath, path.extname(filePath));
  const lastSeparator = baseName.lastIndexOf(' - ');
  if (lastSeparator === -1) {
    return {
      baseName,
      prefix: baseName,
      suffix: baseName
    };
  }
  return {
    baseName,
    prefix: baseName.slice(0, lastSeparator),
    suffix: baseName.slice(lastSeparator + 3)
  };
}

async function findRelatedWorkbooks(seedPath) {
  const info = parseWorkbookName(seedPath);
  const directories = new Set();
  const seedDir = path.dirname(seedPath);
  if (seedDir) directories.add(seedDir);
  if (SPLIT_WORKBOOKS_DIR) {
    try {
      directories.add(path.resolve(SPLIT_WORKBOOKS_DIR));
    } catch (_err) {
      // ignore resolution errors
    }
  }

  const results = new Map();

  for (const dir of directories) {
    if (!dir) continue;
    let exists = false;
    try {
      exists = await pathExists(dir);
    } catch (_err) {
      exists = false;
    }
    if (!exists) continue;

    let entries;
    try {
      entries = await fs.promises.readdir(dir);
    } catch (_err) {
      continue;
    }

    for (const entry of entries) {
      if (!entry.toLowerCase().endsWith('.xlsx')) continue;
      const fullPath = path.join(dir, entry);
      const entryInfo = parseWorkbookName(fullPath);
      if (entryInfo.prefix !== info.prefix) continue;
      const resolved = path.resolve(fullPath);
      if (!results.has(resolved)) {
        results.set(resolved, { path: resolved, info: entryInfo });
      }
    }
  }

  return Array.from(results.values());
}

function formatDisplayDate(dateInput) {
  if (!dateInput) return '';
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return '';
  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).format(date);
}

function formatTimeLabel(start, end) {
  const fmt = (val) => {
    if (!val) return '';
    let s = String(val).trim();
    if (!s) return '';
    s = s.replace(/\./g, ':').replace(/\s+/g, '');
    let mer = null;
    const lower = s.toLowerCase();
    if (/(am|pm)$/.test(lower)) { mer = lower.slice(-2); s = lower.slice(0, -2); }
    let h = 0; let m = 0;
    if (/^\d{1,2}:\d{2}$/.test(s)) { const parts = s.split(':'); h = Number(parts[0]); m = Number(parts[1]); }
    else if (/^\d{3,4}$/.test(s)) { const v = s.padStart(4, '0'); h = Number(v.slice(0,2)); m = Number(v.slice(2)); }
    else if (/^\d{1,2}$/.test(s)) { h = Number(s); m = 0; }
    else { return String(val); }
    if (Number.isNaN(h) || Number.isNaN(m)) return '';
    if (mer) { if (mer === 'pm' && h < 12) h += 12; if (mer === 'am' && h === 12) h = 0; }
    h = Math.max(0, Math.min(23, h)); m = Math.max(0, Math.min(59, m));
    const outMer = h >= 12 ? 'pm' : 'am';
    const h12 = (h % 12) === 0 ? 12 : (h % 12);
    const mm = String(m).padStart(2, '0');
    return `${h12}:${mm} ${outMer}`;
  };
  const a = fmt(start);
  const b = fmt(end);
  if (a && b) return `${a} – ${b}`;
  return a || b || '';
}

async function buildPersonnelLogHtml(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id);
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');

  const business = await db.getBusinessById(businessId);
  if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');

  const fromDate = options.fromDate || options.from_date || new Date().toISOString().slice(0, 10);
  const toDate = options.toDate || options.to_date || null;

  const includeArchived = options.includeArchived === true || options.include_archived === true;

  const all = await db.getAhmenJobsheets({ businessId, includeArchived });

  const isOnOrAfter = (d, base) => {
    try {
      const a = new Date(String(d));
      const b = new Date(String(base));
      if (Number.isNaN(a) || Number.isNaN(b)) return true;
      // Compare dates in local time by yyyy-mm-dd
      const aa = a.toISOString().slice(0,10);
      const bb = b.toISOString().slice(0,10);
      return aa >= bb;
    } catch (_) { return true; }
  };
  const isOnOrBefore = (d, base) => {
    try {
      const a = new Date(String(d));
      const b = new Date(String(base));
      if (Number.isNaN(a) || Number.isNaN(b)) return true;
      const aa = a.toISOString().slice(0,10);
      const bb = b.toISOString().slice(0,10);
      return aa <= bb;
    } catch (_) { return true; }
  };

  const upcoming = (all || []).filter(js => js && js.event_date && (!fromDate || isOnOrAfter(js.event_date, fromDate)) && (!toDate || isOnOrBefore(js.event_date, toDate)));

  upcoming.sort((a, b) => {
    const da = new Date(a.event_date);
    const dbd = new Date(b.event_date);
    const cmp = da - dbd;
    if (cmp !== 0) return cmp;
    const sa = String(a.event_start || '');
    const sb = String(b.event_start || '');
    return sa.localeCompare(sb);
  });

  // Build singer pool map for id->name lookups
  let poolMap = new Map();
  try {
    const pricing = await ahmenCosting.loadPricingConfig();
    const pool = Array.isArray(pricing?.singerPool) ? pricing.singerPool : [];
    poolMap = new Map(pool.map(s => [String(s.id), String(s.name || s.id)]));
  } catch (_) {}

  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  // Column selection
  const allowedColumns = ['date','time','status','client','event','venue','personnel','singer_count','total','notes'];
  const selectedColumns = Array.isArray(options.columns)
    ? options.columns.map(String).map(s => s.toLowerCase()).filter(k => allowedColumns.includes(k))
    : ['date','time','client','event','venue','personnel'];

  const headerLabels = {
    date: 'Date',
    time: 'Time',
    status: 'Status',
    client: 'Client',
    event: 'Event',
    venue: 'Venue',
    personnel: 'Personnel',
    singer_count: 'Singers',
    total: 'Total',
    notes: 'Notes'
  };

  const rowsHtml = upcoming.map(js => {
    let selected = [];
    try {
      if (Array.isArray(js.pricing_selected_singers)) {
        selected = js.pricing_selected_singers;
      } else if (typeof js.pricing_selected_singers === 'string' && js.pricing_selected_singers.trim()) {
        selected = JSON.parse(js.pricing_selected_singers);
      }
    } catch (_) { selected = []; }

    const names = selected.map(entry => {
      try {
        const id = entry && (entry.id ?? entry.singerId ?? entry.value);
        const name = entry && entry.name;
        const label = name || (id != null && poolMap.get(String(id))) || (id != null ? String(id) : '');
        return label || '';
      } catch (_) { return ''; }
    }).filter(Boolean);

    const specialist = String(js.specialist_singers || '').trim();
    const personnel = [
      names.join(', '),
      specialist ? `(Specialist: ${specialist})` : ''
    ].filter(Boolean).join(' ');

    const date = formatDisplayDate(js.event_date);
    const time = formatTimeLabel(js.event_start, js.event_end);
    const status = String(js.status || '').trim();
    const client = js.client_name || '';
    const eventType = js.event_type || '';
    const venue = [js.venue_name || '', js.venue_town || '', js.venue_postcode || ''].filter(Boolean).join(', ');
    const singerCount = names.length;
    const totalNumber = (() => {
      const explicit = Number(js.pricing_total);
      if (Number.isFinite(explicit) && explicit > 0) return explicit;
      const a = Number(js.ahmen_fee) || 0;
      const p = Number(js.production_fees) || 0;
      const sum = a + p;
      return Number.isFinite(sum) ? sum : 0;
    })();
    const total = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(totalNumber);
    const notes = (js.notes || '').toString().trim().slice(0, 180);

    const base = { date, time, status, client, event: eventType, venue, personnel, singer_count: String(singerCount), total, notes };
    const cells = selectedColumns.map(key => `<td>${esc(base[key] || '')}</td>`).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  const title = 'Upcoming Events – Personnel';
  const subtitleParts = [];
  if (fromDate) subtitleParts.push(`from ${formatDisplayDate(fromDate)}`);
  if (toDate) subtitleParts.push(`to ${formatDisplayDate(toDate)}`);
  const subtitle = subtitleParts.join(' ');

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>${esc(title)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #111827; }
    .container { padding: 24px; }
    h1 { font-size: 20px; margin: 0 0 4px; }
    .subtitle { color: #6b7280; font-size: 12px; margin: 0 0 16px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #e5e7eb; padding: 6px 8px; font-size: 12px; vertical-align: top; }
    th { background: #f3f4f6; text-align: left; }
    .c-date { white-space: nowrap; width: 80px; }
    .c-time { white-space: nowrap; width: 90px; color: #374151; }
    .c-client { width: 160px; }
    .c-event { width: 140px; }
    .c-venue { width: 220px; }
    .c-personnel { width: auto; }
    .empty { color: #6b7280; font-size: 13px; padding: 12px 0; }
  </style>
  </head>
  <body>
    <div class="container">
      <h1>${esc(title)}</h1>
      ${subtitle ? `<div class="subtitle">${esc(subtitle)}</div>` : ''}
      ${upcoming.length === 0 ? (`<div class="empty">No upcoming events.</div>`) : (`
        <table>
          <thead>
            <tr>
              ${selectedColumns.map(key => `<th>${esc(headerLabels[key] || key)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      `)}
    </div>
  </body>
  </html>`;

  // Target path: Documents folder under business save_path -> Reports
  const todayIso = new Date().toISOString().slice(0, 10);
  const reportsDir = path.resolve(business.save_path, 'Reports');
  await ensureDirectoryExists(reportsDir);
  const fileName = toDate
    ? `Upcoming Personnel ${fromDate} to ${toDate}.pdf`
    : `Upcoming Personnel ${todayIso}.pdf`;
  const targetPath = path.join(reportsDir, fileName);

  return { html, targetPath };
}

async function buildPersonnelLogText(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id);
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');

  const fromDate = options.fromDate || options.from_date || new Date().toISOString().slice(0, 10);
  const toDate = options.toDate || options.to_date || null;
  const includeArchived = options.includeArchived === true || options.include_archived === true;

  const all = await db.getAhmenJobsheets({ businessId, includeArchived });

  const isOnOrAfter = (d, base) => {
    try { const a = new Date(String(d)); const b = new Date(String(base)); if (Number.isNaN(a) || Number.isNaN(b)) return true; return a.toISOString().slice(0,10) >= b.toISOString().slice(0,10); } catch (_) { return true; }
  };
  const isOnOrBefore = (d, base) => {
    try { const a = new Date(String(d)); const b = new Date(String(base)); if (Number.isNaN(a) || Number.isNaN(b)) return true; return a.toISOString().slice(0,10) <= b.toISOString().slice(0,10); } catch (_) { return true; }
  };

  const upcoming = (all || []).filter(js => js && js.event_date && (!fromDate || isOnOrAfter(js.event_date, fromDate)) && (!toDate || isOnOrBefore(js.event_date, toDate)));
  upcoming.sort((a, b) => { const da = new Date(a.event_date); const dbd = new Date(b.event_date); const cmp = da - dbd; if (cmp !== 0) return cmp; return String(a.event_start||'').localeCompare(String(b.event_start||'')); });

  // Pool lookup
  let poolMap = new Map();
  try {
    const pricing = await ahmenCosting.loadPricingConfig();
    const pool = Array.isArray(pricing?.singerPool) ? pricing.singerPool : [];
    poolMap = new Map(pool.map(s => [String(s.id), String(s.name || s.id)]));
  } catch (_) {}

  const allowedColumns = ['date','time','status','client','event','venue','personnel','singer_count','total','notes'];
  const selectedColumns = Array.isArray(options.columns)
    ? options.columns.map(String).map(s => s.toLowerCase()).filter(k => allowedColumns.includes(k))
    : ['date','time','client','event','venue','personnel'];
  const headerLabels = {
    date: 'Date', time: 'Time', status: 'Status', client: 'Client', event: 'Event', venue: 'Venue', personnel: 'Personnel', singer_count: 'Singers', total: 'Total', notes: 'Notes'
  };

  const formatMoney = (n) => new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(n || 0);
  const singleLine = options.singleLine !== false; // default true for WhatsApp
  const bullet = options.bullet !== false; // default true

  const lines = upcoming.map(js => {
    let selected = [];
    try {
      if (Array.isArray(js.pricing_selected_singers)) selected = js.pricing_selected_singers;
      else if (typeof js.pricing_selected_singers === 'string' && js.pricing_selected_singers.trim()) selected = JSON.parse(js.pricing_selected_singers);
    } catch (_) { selected = []; }
    const names = selected.map(entry => {
      const id = entry && (entry.id ?? entry.singerId ?? entry.value);
      const name = entry && entry.name;
      const label = name || (id != null && poolMap.get(String(id))) || (id != null ? String(id) : '');
      return label || '';
    }).filter(Boolean);
    const specialist = String(js.specialist_singers || '').trim();
    const personnel = [names.join(', '), specialist ? `(Specialist: ${specialist})` : ''].filter(Boolean).join(' ');
    const date = formatDisplayDate(js.event_date);
    const time = formatTimeLabel(js.event_start, js.event_end);
    const status = String(js.status || '').trim();
    const client = js.client_name || '';
    const eventType = js.event_type || '';
    const venue = [js.venue_name || '', js.venue_town || '', js.venue_postcode || ''].filter(Boolean).join(', ');
    const singerCount = String(names.length);
    const total = formatMoney((Number(js.pricing_total) && Number(js.pricing_total) > 0) ? Number(js.pricing_total) : (Number(js.ahmen_fee)||0) + (Number(js.production_fees)||0));
    const notes = (js.notes || '').toString().trim().replace(/[\r\n]+/g, ' ').slice(0, 180);

    const base = { date, time, status, client, event: eventType, venue, personnel, singer_count: singerCount, total, notes };
    if (singleLine) {
      const parts = selectedColumns.map(key => base[key]).filter(Boolean);
      const line = parts.join(' — ');
      return `${bullet ? '• ' : ''}${line}`;
    }
    const rows = selectedColumns.map(key => `${headerLabels[key]}: ${base[key] || ''}`);
    return `${bullet ? '• ' : ''}${rows.shift() || ''}\n${rows.join('\n')}`;
  });

  const title = 'Upcoming Events — Personnel';
  const subtitleParts = [];
  if (fromDate) subtitleParts.push(`from ${formatDisplayDate(fromDate)}`);
  if (toDate) subtitleParts.push(`to ${formatDisplayDate(toDate)}`);
  const subtitle = subtitleParts.join(' ');
  const header = [title, subtitle].filter(Boolean).join(' ');
  const text = [header, '', ...lines].join('\n');
  return { text };
}

function buildOutputDirectory(business, context, payload, fileLabel) {
  const baseSavePath = business?.save_path;
  if (!baseSavePath) {
    throw new Error('Configure a documents folder for this business before generating documents.');
  }

  const eventDate = context.event?.event_date || context.jobsheet?.event_date || '';
  const formattedDate = formatDateISO(eventDate);
  const clientName = sanitizeFilenameSegment(context.client?.name || context.jobsheet?.client_name || '');
  const displayDate = sanitizeFilenameSegment(formatDisplayDate(eventDate));

  const folderParts = [formattedDate, clientName, displayDate].filter(Boolean);
  const folderBase = sanitizeFilenameSegment(folderParts.join(' - ') || fileLabel || 'Jobsheet');

  const segments = [baseSavePath];
  segments.push(folderBase);

  return path.join(...segments);
}

function buildFileName(context, payload, definition) {
  const eventDate = context.event?.event_date || context.jobsheet?.event_date || '';
  const formattedDate = formatDateISO(eventDate);
  const clientName = sanitizeFilenameSegment(context.client?.name || context.jobsheet?.client_name || '');
  const displayDate = sanitizeFilenameSegment(formatDisplayDate(eventDate));
  const definitionLabel = sanitizeFilenameSegment(definition?.label || definition?.key || 'Workbook');
  const ext = '.xlsx';

  const folderBase = [formattedDate, clientName, displayDate]
    .map(part => sanitizeFilenameSegment(part))
    .filter(Boolean)
    .join(' - ') || sanitizeFilenameSegment(definitionLabel) || 'Document';

  const baseWithLabel = [folderBase, definitionLabel]
    .map(part => sanitizeFilenameSegment(part))
    .filter(Boolean)
    .join(' - ');

  return {
    folderName: folderBase,
    fileName: `${baseWithLabel}${ext}`
  };
}

async function createWorkbookDocument(payload = {}) {

  const businessId = Number(payload.business_id);
  if (!Number.isInteger(businessId)) {
    throw new Error('business_id is required to generate documents.');
  }

  const templatePath = resolvePath(payload.template_path);
  await ensureFileAccessible(templatePath);

  const business = await db.getBusinessById(businessId);
  if (!business) {
    throw new Error('Business record not found.');
  }

  const context = buildContext(payload, business);

  const definitionKey = payload.definition_key || 'workbook';
  const definition = await db.getDocumentDefinition(businessId, definitionKey);
  const naming = buildFileName(context, payload, definition);
  const directory = buildOutputDirectory(business, context, payload, naming.folderName);
  await fs.promises.mkdir(directory, { recursive: true });

  const targetPath = path.join(directory, naming.fileName);
  // Treat existing workbooks as immutable: do not overwrite
  if (await pathExists(targetPath)) {
    try {
      const existing = await db.getDocumentByFilePath(businessId, targetPath);
      if (existing && existing.is_locked) {
        throw new Error('Workbook is locked');
      }
    } catch (_err) {}
    throw new Error('Workbook already exists');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  // Add a manual page break after row 45 on booking schedule without altering page setup
  try {
    const ws = workbook.getWorksheet('Booking Schedule');
    if (ws) {
      try {
        const row = ws.getRow(45);
        if (row) row.addPageBreak();
      } catch (_) {}
    }
  } catch (_err) {
    // Ignore errors; rely on template defaults
  }

  const bindings = await db.getMergeFieldBindingsByTemplate(TEMPLATE_BINDING_KEY);
  const mergeFields = await db.getMergeFields();
  const placeholderMap = new Map();
  (mergeFields || []).forEach(f => {
    const fieldKey = (f.field_key || '').toLowerCase();
    const placeholder = (f.placeholder || '').toLowerCase();
    const fieldSlug = normalizeTokenKey(fieldKey);
    const placeholderSlug = normalizeTokenKey(placeholder);
    if (fieldKey) placeholderMap.set(fieldKey, f.field_key);
    if (placeholder) placeholderMap.set(placeholder, f.field_key);
    if (fieldSlug) placeholderMap.set(fieldSlug, f.field_key);
    if (placeholderSlug) placeholderMap.set(placeholderSlug, f.field_key);
  });

  const placeholderKeys = collectPlaceholderKeys(workbook);
  const fieldKeySet = new Set((bindings || []).map(binding => binding.field_key).filter(Boolean));
  placeholderKeys.forEach(key => {
    const mapped = placeholderMap.get(String(key).toLowerCase()) || key;
    fieldKeySet.add(mapped);
  });
  const valueSources = await db.getMergeFieldValueSources(Array.from(fieldKeySet)) || {};

  await fillWorkbook(workbook, bindings, valueSources, context);
  replaceWorkbookPlaceholders(workbook, valueSources, context, placeholderMap);
  workbook.calcProperties = workbook.calcProperties || {};
  workbook.calcProperties.fullCalcOnLoad = true;
  sanitizeWorkbookValues(workbook);
  await workbook.xlsx.writeFile(targetPath);

  const clientName = context.client?.name || context.jobsheet?.client_name || null;
  const eventName = context.event?.event_name || context.jobsheet?.event_type || null;
  const eventDate = context.event?.event_date || context.jobsheet?.event_date || null;

  const inserted = await db.addDocument({
    business_id: businessId,
    jobsheet_id: payload.jobsheet_id || null,
    doc_type: 'workbook',
    total_amount: payload.total_amount ?? null,
    balance_due: payload.balance_amount ?? payload.balance_due ?? null,
    due_date: payload.balance_due_date ?? payload.due_date ?? null,
    file_path: targetPath,
    client_name: clientName,
    event_name: eventName,
    event_date: eventDate,
    document_date: payload.document_date || new Date().toISOString(),
    definition_key: payload.definition_key || 'workbook',
    invoice_variant: payload.invoice_variant || null,
    status: 'generated'
  });

  return {
    ok: true,
    file_path: targetPath,
    document_id: inserted?.id || null,
    number: inserted?.number ?? null,
    additional_outputs: []
  };
}

async function createDocument(payload = {}) {
  return createWorkbookDocument(payload);
}

// Build minimal HTML for MCMS invoices/quotes (no Excel dependency)
async function buildMCMSDocumentHtml(options = {}) {
  const businessId = Number(options.business_id ?? options.businessId);
  if (!Number.isInteger(businessId)) throw new Error('business_id is required');
  const rawType = String(options.doc_type || options.type || '').toLowerCase();
  if (!rawType || (rawType !== 'invoice' && rawType !== 'quote')) throw new Error('doc_type must be invoice or quote');

  const business = await db.getBusinessById(businessId);
  if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');

  const client = options.client || options.client_override || {};
  const itemInput = Array.isArray(options.line_items || options.items) ? (options.line_items || options.items) : [];
  const items = itemInput.map((it, idx) => {
    const type = String(it?.item_type || it?.type || '').toLowerCase() || '';
    const desc = (it?.description || '').toString();
    const qty = Number(it?.quantity);
    const unit = (it?.unit || (type === 'studio' ? 'hours' : 'unit')).toString();
    const rate = Number(it?.rate);
    const amount = Number.isFinite(Number(it?.amount)) ? Number(it?.amount) : (Number.isFinite(qty) && Number.isFinite(rate) ? qty * rate : 0);
    return { item_type: type || 'custom', description: desc, quantity: Number.isFinite(qty) ? qty : null, unit, rate: Number.isFinite(rate) ? rate : null, amount, sort_order: idx };
  }).filter(x => (x.amount != null && Number.isFinite(x.amount) && x.amount !== 0) || (x.description && x.description.trim()));
  const computedSubtotal = items.reduce((sum, it) => sum + (Number.isFinite(it.amount) ? it.amount : 0), 0);
  const totalAmount = options.total_amount != null && items.length === 0 ? Number(options.total_amount) : computedSubtotal;
  const dueDate = options.due_date || null;
  const issueDate = options.document_date || new Date().toISOString().slice(0, 10);

  const payload = {
    business_id: businessId,
    client_override: client,
    event_override: { event_date: issueDate },
    total_amount: totalAmount,
    due_date: dueDate,
    document_date: new Date().toISOString()
  };
  const context = buildContext(payload, business);
  const docLabel = rawType === 'invoice' ? 'Invoice' : 'Quote';
  const naming = buildFileName(context, payload, { label: docLabel, key: `${rawType}_html` });
  const directory = buildOutputDirectory(business, context, payload, naming.folderName);
  await ensureDirectoryExists(directory);

  // Reserve a document number
  const inserted = await db.addDocument({
    business_id: businessId,
    doc_type: rawType,
    status: 'draft',
    total_amount: totalAmount,
    balance_due: totalAmount,
    due_date: dueDate,
    client_name: context.client?.name || null,
    event_name: null,
    event_date: null,
    document_date: payload.document_date,
    definition_key: `${rawType}_html`
  });
  const number = inserted?.number != null ? Number(inserted.number) : null;
  try { if (items.length) await db.saveDocumentItems(inserted.id, items); } catch (_) {}

  const baseName = `${naming.folderName} - ${docLabel}`;
  const suffixCode = rawType === 'invoice' ? (number != null ? `INV-${number}` : 'INV') : (number != null ? `Q-${number}` : 'Q');
  let targetPath = path.join(directory, `${baseName} (${suffixCode}).pdf`);
  let n = 2;
  // Version file if exists
  while (await pathExists(targetPath)) {
    targetPath = path.join(directory, `${baseName} (${suffixCode}) (${n}).pdf`);
    n += 1;
    if (n > 1000) break;
  }

  const todayDisplay = formatDisplayDate(issueDate);
  const dueDisplay = formatDisplayDate(dueDate);
  const amountDisplay = formatCurrencyGBP(totalAmount);
  const clientAddressLines = [client.address1 || client.address, client.address2, client.town, client.postcode].filter(Boolean).join('<br>');

  const rowsHtml = (items.length ? items : [{ description: (rawType === 'invoice' ? 'Services rendered' : 'Quoted services'), quantity: null, unit: '', rate: null, amount: totalAmount }])
    .map(it => {
      const qty = it.quantity != null && Number.isFinite(it.quantity) ? it.quantity : '';
      const rate = it.rate != null && Number.isFinite(it.rate) ? formatCurrencyGBP(it.rate) : '';
      const lineTotal = it.amount != null && Number.isFinite(it.amount) ? formatCurrencyGBP(it.amount) : '';
      return `<tr><td>${escapeHtml(it.description || '')}</td><td class="right">${qty}</td><td>${escapeHtml(it.unit || '')}</td><td class="right">${rate}</td><td class="right">${lineTotal}</td></tr>`;
    }).join('');

  // Use inline HTML if provided, otherwise custom HTML template if saved in settings
  let htmlTemplate = (options && typeof options.inline_html === 'string') ? options.inline_html : '';
  try {
    if (!htmlTemplate) {
      const tmpl = await module.exports.getHtmlTemplate({ businessId: businessId, docType: rawType });
      if (tmpl && typeof tmpl.html === 'string') htmlTemplate = tmpl.html;
    }
  } catch (_) {}

  let html = '';
  if (htmlTemplate) {
    const tokens = {
      business_name: business.business_name || 'MCMS',
      client_name: context.client?.name || '',
      client_address_html: clientAddressLines || '',
      invoice_title: docLabel,
      invoice_code: suffixCode,
      issue_date: todayDisplay || '',
      due_date: dueDisplay || '',
      total_amount: amountDisplay,
      items_table: buildItemsTableHtml(items.length ? items : [{ description: (rawType === 'invoice' ? 'Services rendered' : 'Quoted services'), quantity: null, unit: '', rate: null, amount: totalAmount }])
    };
    html = renderTemplateHtml(htmlTemplate, tokens);
  } else {
    html = `<!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <title>${docLabel}${number != null ? ` #${number}` : ''}</title>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #111827; margin: 40px; }
      h1 { font-size: 22px; margin: 0 0 4px 0; }
      .muted { color: #6b7280; }
      .row { display: flex; justify-content: space-between; align-items: flex-start; }
      .section { margin: 18px 0; }
      table { width: 100%; border-collapse: collapse; margin-top: 12px; }
      th, td { padding: 10px 8px; border-bottom: 1px solid #e5e7eb; text-align: left; }
      th { background: #f9fafb; font-weight: 600; font-size: 12px; }
      .right { text-align: right; }
      .total-row td { border-top: 2px solid #111827; font-weight: 700; }
      .footer { margin-top: 28px; font-size: 12px; color: #6b7280; }
    </style>
  </head>
  <body>
    <div class="row">
      <div>
        <h1>${business.business_name || 'MCMS'}</h1>
        <div class="muted">${options.fromEmail || 'mottitemp@hotmail.com'}</div>
      </div>
      <div style="text-align:right">
        <div style="font-size:28px; font-weight:700;">${docLabel}</div>
        <div class="muted">${suffixCode}${number == null ? '' : ''}</div>
      </div>
    </div>

    <div class="section row">
      <div>
        <div style="font-weight:600;">Bill to</div>
        <div>${context.client?.name || ''}</div>
        <div class="muted">${clientAddressLines || ''}</div>
      </div>
      <div style="text-align:right">
        <div><span class="muted">Date:</span> ${todayDisplay || ''}</div>
        ${rawType === 'invoice' ? `<div><span class="muted">Due:</span> ${dueDisplay || ''}</div>` : ''}
      </div>
    </div>

    <table>
      <thead>
        <tr>
          <th>Description</th>
          <th class="right" style="width:80px">Qty</th>
          <th style="width:80px">Unit</th>
          <th class="right" style="width:120px">Rate</th>
          <th class="right" style="width:120px">Line Total</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
        <tr class="total-row"><td colspan="4">Total</td><td class="right">${formatCurrencyGBP(totalAmount)}</td></tr>
      </tbody>
    </table>

    <div class="footer">
      ${rawType === 'invoice' ? 'Please make payment by the due date.' : 'This quote is provided for your consideration.'}
    </div>
  </body>
  </html>`;
  }

  return { html, targetPath, number, document_id: inserted?.id || null, business_id: businessId, doc_type: rawType };
}

const documentWatchers = new Map();
const watcherCallbacks = new Map();
const watcherTimers = new Map();

async function watchDocumentsFolder(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required to watch documents folder.');
  }

  const business = await db.getBusinessById(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured for this business.');
  }

  const rootPath = path.resolve(business.save_path);
  await ensureDirectoryExists(rootPath);

  if (typeof options.onChange === 'function') {
    watcherCallbacks.set(businessId, options.onChange);
  }

  if (documentWatchers.has(businessId)) {
    return { ok: true, watching: true };
  }

  const triggerChange = async () => {
    try {
      const callback = watcherCallbacks.get(businessId);
      if (callback) {
        callback({ businessId });
      }
    } catch (err) {
      console.error('Failed to notify documents change', err);
    }
  };

  let watcher;
  if (chokidar) {
    watcher = chokidar.watch(rootPath, {
      ignoreInitial: true,
      persistent: true,
      awaitWriteFinish: { stabilityThreshold: 200, pollInterval: 50 },
      depth: 10
    });
    const schedule = () => {
      const existingTimer = watcherTimers.get(businessId);
      if (existingTimer) clearTimeout(existingTimer);
      const timer = setTimeout(triggerChange, 250);
      watcherTimers.set(businessId, timer);
    };
    ['add', 'addDir', 'change', 'unlink', 'unlinkDir'].forEach(evt => watcher.on(evt, schedule));
    watcher.on('error', err => console.error('Documents watcher error', err));
  } else {
    watcher = fs.watch(rootPath, { recursive: true }, () => {
      const existingTimer = watcherTimers.get(businessId);
      if (existingTimer) {
        clearTimeout(existingTimer);
      }
      const timer = setTimeout(triggerChange, 300);
      watcherTimers.set(businessId, timer);
    });
    watcher.on('error', (err) => {
      console.error('Documents watcher error', err);
    });
  }

  documentWatchers.set(businessId, watcher);
  return { ok: true, watching: true };
}

function unwatchDocumentsFolder(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    return { ok: true, watching: false };
  }

  const watcher = documentWatchers.get(businessId);
  if (watcher) {
    try {
      watcher.close();
    } catch (err) {
      console.warn('Failed to close documents watcher', err);
    }
  }
  documentWatchers.delete(businessId);

  const timer = watcherTimers.get(businessId);
  if (timer) {
    clearTimeout(timer);
    watcherTimers.delete(businessId);
  }
  watcherCallbacks.delete(businessId);

  return { ok: true, watching: false };
}

async function filterDocumentsByExistingFiles(documents, options = {}) {
  if (!Array.isArray(documents)) return [];
  const includeMissing = options.includeMissing === true;

  const enriched = await Promise.all(documents.map(async (doc) => {
    const filePath = doc?.file_path || doc?.filePath;
    const fileAvailable = filePath ? await pathExists(filePath) : false;
    return { ...doc, file_available: fileAvailable };
  }));

  return includeMissing ? enriched : enriched.filter(doc => doc.file_available);
}


async function listJobsheetDocuments(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required to list jobsheet documents.');
  }

  const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
  const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;

  const documents = await db.getDocuments({ businessId });
  const enriched = await filterDocumentsByExistingFiles(documents, { includeMissing: false });

  // Deduplicate by exact file_path (keep the latest document_id) and clean DB for duplicates
  try {
    const byPath = new Map();
    for (const doc of enriched) {
      const fp = doc?.file_path;
      if (!fp) continue;
      const existing = byPath.get(fp);
      if (!existing || Number(doc.document_id) > Number(existing.document_id)) {
        byPath.set(fp, doc);
      }
    }
    const dups = [];
    for (const doc of enriched) {
      const fp = doc?.file_path;
      if (!fp) continue;
      const primary = byPath.get(fp);
      if (primary && Number(primary.document_id) !== Number(doc.document_id)) {
        dups.push(doc);
      }
    }
    // Clear file_path for duplicates in DB so they stop showing up in future
    await Promise.all(dups.map(doc => db.setDocumentFilePath(doc.document_id, null).catch(() => null)));
    // Rebuild enriched without duplicates
    const deduped = Array.from(byPath.values());
    // Replace enriched reference
    enriched.length = 0;
    deduped.forEach(doc => enriched.push(doc));
  } catch (_err) {
    // If any cleanup fails, continue with current list
  }

  // DB hygiene: clear file_path for any entries whose file no longer exists
  try {
    const missing = Array.isArray(documents) ? documents.filter(d => d?.file_path && !enriched.some(e => e?.file_path === d.file_path)) : [];
    await Promise.all(missing.map(d => db.clearDocumentPath(businessId, d.file_path).catch(() => null)));
  } catch (_err) {}

  const folderSet = new Set();

  const mapped = enriched.map((doc) => {
    const filePath = doc?.file_path || null;
    const fileName = filePath ? path.basename(filePath) : null;
    const folderPath = filePath ? path.dirname(filePath) : null;
    if (folderPath) folderSet.add(folderPath);

    let parsed = null;
    if (filePath) {
      try {
        parsed = parseWorkbookName(filePath);
      } catch (_err) {
        parsed = null;
      }
    }

    const label = doc?.definition_label
      || doc?.label
      || fileName
      || doc?.doc_type
      || 'Document';

    return {
      ...doc,
      file_name: fileName,
      file_prefix: null,
      file_suffix: null,
      display_label: label,
      folder_path: folderPath
    };
  });

  let jobsheetFolder = null;
  try {
    const business = await db.getBusinessById(businessId);
    if (business?.save_path) {
      const payload = {
        business_id: businessId,
        jobsheet_id: jobsheetId,
        jobsheet_snapshot: options.jobsheetSnapshot && typeof options.jobsheetSnapshot === 'object'
          ? { ...options.jobsheetSnapshot }
          : {},
        client_override: options.clientOverride && typeof options.clientOverride === 'object'
          ? { ...options.clientOverride }
          : {},
        event_override: options.eventOverride && typeof options.eventOverride === 'object'
          ? { ...options.eventOverride }
          : {},
        pricing_snapshot: options.pricingSnapshot && typeof options.pricingSnapshot === 'object'
          ? { ...options.pricingSnapshot }
          : {}
      };
      const context = buildContext(payload, business);
      try {
        jobsheetFolder = buildOutputDirectory(business, context, payload, 'Documents');
      } catch (_err) {
        jobsheetFolder = null;
      }
    }
  } catch (_err) {
    jobsheetFolder = null;
  }

  const folders = Array.from(folderSet);

  return {
    documents: mapped,
    folders,
    jobsheet_folder: jobsheetFolder
  };
}


async function exportWorkbookPdfs(options = {}) {
  const providedPath = options.filePath || options.file_path;
  if (!providedPath || typeof providedPath !== 'string') {
    throw new Error('filePath is required to export PDFs.');
  }

  const normalizedPath = path.resolve(providedPath.trim());
  await ensureFileAccessible(normalizedPath);

  const businessId = options.businessId != null ? Number(options.businessId) : null;

  const masterInfo = parseWorkbookName(normalizedPath);
  const jobDirectory = path.dirname(normalizedPath);

  const includeRelated = options.includeRelated === true;
  let workbooks = [{ path: normalizedPath, info: masterInfo }];
  if (includeRelated) {
    const relatedWorkbooks = await findRelatedWorkbooks(normalizedPath);
    relatedWorkbooks.forEach(entry => {
      if (!entry || !entry.path) return;
      if (path.resolve(entry.path) === path.resolve(normalizedPath)) return;
      workbooks.push(entry);
    });
  }

  const outputs = [];

  const activeSheetOnly = options.activeSheetOnly === true;

  for (const entry of workbooks) {
    const workbookPath = entry.path;
    const info = entry.info || parseWorkbookName(workbookPath);
    const targetPdfName = `${info.baseName}.pdf`;
    const targetPdfPath = path.join(jobDirectory, targetPdfName);
    let effectivePdfPath = targetPdfPath;

    try {
      // Determine if this is an invoice export up front to allow safe versioning
      let variantForVersioning = null;
      if (businessId != null) {
        try {
          const wbDoc = await db.getDocumentByFilePath(businessId, workbookPath);
          const defKey = wbDoc?.definition_key || null;
          if (defKey) {
            try {
              const def = await db.getDocumentDefinition(businessId, defKey);
              const v = (def?.invoice_variant || '').toLowerCase();
              if (v === 'deposit' || v === 'balance') variantForVersioning = v;
            } catch (_) {}
          }
        } catch (_) {}
      }
      // Fallback: infer from file suffix if definition missing
      if (!variantForVersioning && info && info.suffix) {
        const sfx = String(info.suffix).toLowerCase();
        if (sfx.includes('deposit')) variantForVersioning = 'deposit';
        else if (sfx.includes('balance')) variantForVersioning = 'balance';
      }

      // Treat existing PDFs as immutable: do not overwrite; version invoices instead
      try {
        const exists = await pathExists(targetPdfPath);
        if (exists) {
          // If DB has a locked record, report locked
          if (businessId != null) {
            try {
              const existing = await db.getDocumentByFilePath(businessId, targetPdfPath);
              if (existing && existing.is_locked) {
                outputs.push({ success: false, sheet: info.suffix, error: 'PDF is locked' });
                continue;
              }
            } catch (_) {}
          }
          if (variantForVersioning) {
            // Create a versioned filename: "name (2).pdf", "name (3).pdf", ...
            const base = path.basename(targetPdfPath, '.pdf');
            const dir = path.dirname(targetPdfPath);
            let n = 2;
            // eslint-disable-next-line no-constant-condition
            while (true) {
              const candidate = path.join(dir, `${base} (${n}).pdf`);
              // eslint-disable-next-line no-await-in-loop
              const taken = await pathExists(candidate);
              if (!taken) { effectivePdfPath = candidate; break; }
              n += 1;
              if (n > 1000) { break; }
            }
          } else {
            outputs.push({ success: false, sheet: info.suffix, error: 'PDF already exists' });
            continue;
          }
        }
      } catch (_err) {}

      // If a PDF already exists and is locked in DB, block export (defensive)
      if (businessId != null) {
        try {
          const existing = await db.getDocumentByFilePath(businessId, effectivePdfPath);
          if (existing && existing.is_locked) {
            outputs.push({ success: false, sheet: info.suffix, error: 'PDF is locked' });
            continue;
          }
        } catch (_err) {}
      }

      // Determine if this workbook corresponds to an invoice (deposit/balance) for stamping + numbering
      let stampedLabel = 'PDF';
      let performedStamp = false;
      let createdInvoiceId = null;

      if (businessId != null) {
        try {
          const wbDoc = await db.getDocumentByFilePath(businessId, workbookPath);
          const definitionKey = wbDoc?.definition_key || null;
          const jobsheetId = wbDoc?.jobsheet_id != null ? Number(wbDoc.jobsheet_id) : null;
          let variant = null;
          if (definitionKey) {
            try {
              const def = await db.getDocumentDefinition(businessId, definitionKey);
              const v = (def?.invoice_variant || '').toLowerCase();
              if (v === 'deposit' || v === 'balance') variant = v;
            } catch (_err) {}
          }
          // Fallback: infer from workbook file name suffix if definition is missing
          if (!variant && info && info.suffix) {
            const s = String(info.suffix).toLowerCase();
            if (s.includes('deposit')) variant = 'deposit';
            else if (s.includes('balance')) variant = 'balance';
          }

          if (variant) {
            // Ensure environment is ready before reserving an invoice number (Excel only, no Finder fallback)
            try {
              const pf = await preflightPdfExport({ filePath: workbookPath, excelOnly: true });
              if (!pf?.ok) {
                throw new Error(pf?.message || 'Preflight failed');
              }
            } catch (pfErr) {
              throw pfErr;
            }
            // Reserve the next invoice number by creating the DB row first (so counter increments only when success, we’ll roll back on failure)
            const clientName = wbDoc?.client_name || null;
            const eventName = wbDoc?.event_name || null;
            const eventDate = wbDoc?.event_date || null;
            const documentDate = new Date().toISOString();
            const totalAmount = wbDoc?.total_amount ?? null;
            const balanceDue = wbDoc?.balance_due ?? null;
            const dueDate = wbDoc?.due_date ?? null;

            // Insert invoice row with auto-assigned number, without file_path yet
            const inserted = await db.addDocument({
              business_id: businessId,
              jobsheet_id: jobsheetId,
              doc_type: 'invoice',
              number: options && Number.isInteger(options.requestedNumber) && options.requestedNumber > 0 ? Number(options.requestedNumber) : undefined,
              status: 'issued',
              total_amount: totalAmount,
              balance_due: balanceDue,
              due_date: dueDate,
              file_path: null,
              client_name: clientName,
              event_name: eventName,
              event_date: eventDate,
              document_date: documentDate,
              definition_key: definitionKey,
              invoice_variant: variant
            });

            const invoiceNumber = inserted?.number != null ? Number(inserted.number) : null;
            createdInvoiceId = inserted?.id || null;

            // Determine prefix and stamp
            const prefix = await readInvoicePrefixFromWorkbook(workbookPath);
            const stampText = `${prefix}${invoiceNumber != null ? invoiceNumber : ''}`;

            // Build final numbered filename: base name + " (INV-###).pdf" (with versioning if necessary)
            try {
              if (invoiceNumber != null) {
                const invBase = `${info.baseName} (INV-${invoiceNumber})`;
                let candidate = path.join(jobDirectory, `${invBase}.pdf`);
                let n = 2;
                // eslint-disable-next-line no-constant-condition
                while (await pathExists(candidate)) {
                  candidate = path.join(jobDirectory, `${invBase} (${n}).pdf`);
                  n += 1;
                  if (n > 1000) break;
                }
                effectivePdfPath = candidate;
              }
            } catch (_) {}
            // Stamp directly via AppleScript; target the expected invoice sheet
            const stampSheetName = variant === 'deposit' ? 'Invoice – Deposit' : 'Invoice – Balance';
            await saveWorkbookAsPdf(workbookPath, effectivePdfPath, {
              activeSheetOnly,
              stampCell: 'E9',
              stampText,
              stampSheetName,
              stampVariant: variant
            });
            performedStamp = true;
            stampedLabel = invoiceNumber != null ? `Invoice #${invoiceNumber}` : 'Invoice PDF';

            // Update the invoice record with final file path and reminder date per variant
            let reminderDate = null;
            if (variant === 'balance' && jobsheetId != null) {
              try {
                const js = await db.getAhmenJobsheet(jobsheetId);
                reminderDate = js?.balance_reminder_date || null;
              } catch (_err) {}
            }
            await db.updateDocumentStatus(createdInvoiceId, {
              file_path: effectivePdfPath,
              status: 'issued',
              reminder_date: reminderDate,
              due_date: dueDate,
              balance_due: balanceDue,
              total_amount: totalAmount
            });
          }
        } catch (_err) {
          // Do not export unnumbered invoice PDFs; record failure and continue
          outputs.push({ success: false, sheet: info.suffix, error: _err?.message || 'Invoice export failed' });
          continue;
        }
      }

      if (!performedStamp) {
        await saveWorkbookAsPdf(workbookPath, effectivePdfPath, { activeSheetOnly });
      }

      outputs.push({
        success: true,
        sheet: info.suffix,
        label: stampedLabel,
        file_path: effectivePdfPath
      });
    } catch (err) {
      // If export errored but the PDF exists (e.g., macOS prompt timing), salvage success and keep invoice row
      try {
        const existsAfterFail = await pathExists(effectivePdfPath);
        if (existsAfterFail) {
          if (businessId != null && (typeof createdInvoiceId === 'number' || (createdInvoiceId != null && Number.isInteger(Number(createdInvoiceId))))) {
            try {
              await db.updateDocumentStatus(createdInvoiceId, { file_path: effectivePdfPath, status: 'issued' });
            } catch (_) {}
          }
          outputs.push({ success: true, sheet: info.suffix, label: stampedLabel || 'Invoice PDF', file_path: effectivePdfPath });
        } else {
          // No file; roll back the invoice row and counter if we reserved one
          try {
            if (businessId != null && (typeof createdInvoiceId === 'number' || (createdInvoiceId != null && Number.isInteger(Number(createdInvoiceId))))) {
              try { await db.deleteDocument(createdInvoiceId); } catch (_err) {}
              try {
                const maxNum = await db.getMaxInvoiceNumber(businessId);
                const last = Number.isInteger(Number(maxNum)) ? Number(maxNum) : 0;
                await db.setLastInvoiceNumber(businessId, last);
              } catch (_err) {}
            }
          } catch (_) {}
          outputs.push({ success: false, sheet: info.suffix, error: err?.message || 'Unable to export sheet' });
        }
      } catch (_checkErr) {
        outputs.push({ success: false, sheet: info.suffix, error: err?.message || 'Unable to export sheet' });
      }
    }
  }

  const ok = outputs.some(item => item.success);
  let message = '';
  if (!ok) {
    const firstError = outputs.find(o => o && o.success === false && o.error);
    if (firstError && firstError.error) {
      message = `Export failed: ${firstError.error}`;
    } else if (!outputs.length) {
      message = 'Export failed: no sheets processed.';
    } else {
      message = 'Export failed.';
    }
  }
  try {
    if (businessId != null && ok) {
      const cb = watcherCallbacks.get(businessId);
      if (typeof cb === 'function') {
        cb({ businessId });
      }
    }
  } catch (_err) {}
  return { ok, workbook_path: normalizedPath, outputs, message };
}

function formatGigDate(value) {
  if (!value) return '';
  try {
    const d = new Date(value);
    if (Number.isNaN(d.valueOf())) return String(value);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch (_) { return String(value); }
}

function safeHtml(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

async function buildGigInfoHtml(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id);
  const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');
  if (!Number.isInteger(jobsheetId)) throw new Error('jobsheetId is required');

  const js = await db.getAhmenJobsheet(jobsheetId);
  if (!js) throw new Error('Jobsheet not found');

  let info = {};
  if (options && options.gigInfo && typeof options.gigInfo === 'object') {
    info = {
      values: (options.gigInfo.values && typeof options.gigInfo.values === 'object') ? options.gigInfo.values : {},
      include: (options.gigInfo.include && typeof options.gigInfo.include === 'object') ? options.gigInfo.include : {}
    };
  } else {
    try { info = js.gig_info ? JSON.parse(js.gig_info) || {} : {}; } catch (_) { info = {}; }
  }
  const values = (info && info.values) || {};
  const include = (info && info.include) || {};
  const compact = include.compact_spacing === true;

  const titleDate = include.event_date !== false ? formatGigDate(js.event_date) : '';
  const header = `Gig info sheet${titleDate ? `: ${titleDate}` : ''}`;

  const clientName = include.client_name !== false ? (values.client_name || js.client_name || '') : '';
  const eventType = include.event_type !== false ? (values.event_type || js.event_type || '') : '';

  const parts = [];
  if (clientName) parts.push(`<div class="section"><div class="label">Client</div><div class="value">${safeHtml(clientName)}</div></div>`);
  if (eventType) parts.push(`<div class="section"><div class="label">Event</div><div class="value">${safeHtml(eventType)}</div></div>`);

  // Venue block
  const venueIncluded = include.venue_block !== false;
  if (venueIncluded) {
    const vName = values.venue_name || js.venue_name || '';
    const a1 = values.venue_address1 || js.venue_address1 || '';
    const a2 = values.venue_address2 || js.venue_address2 || '';
    const a3 = values.venue_address3 || js.venue_address3 || '';
    const town = values.venue_town || js.venue_town || '';
    const pc = values.venue_postcode || js.venue_postcode || '';
    const lines = [vName, a1, a2, a3, [town, pc].filter(Boolean).join(' ')].filter(Boolean);
    if (lines.length) {
      parts.push(`<div class="section"><div class="label">Venue</div><div class="value">${lines.map(safeHtml).join('<br>')}</div></div>`);
    }
  }

  // Schedule (default to included unless explicitly disabled). Include event times (12-hour am/pm) and call time 1:15 before start.
  {
    const formatTime = (input) => {
      if (!input) return '';
      let s = String(input).trim();
      if (!s) return '';
      s = s.replace(/\./g, ':').replace(/\s+/g, '');
      let mer = null;
      const lower = s.toLowerCase();
      if (/(am|pm)$/.test(lower)) {
        mer = lower.slice(-2);
        s = lower.slice(0, -2);
      }
      let h = 0; let m = 0;
      if (/^\d{1,2}:\d{2}$/.test(s)) {
        const parts = s.split(':');
        h = Number(parts[0]);
        m = Number(parts[1]);
      } else if (/^\d{3,4}$/.test(s)) {
        const v = s.padStart(4, '0');
        h = Number(v.slice(0, 2));
        m = Number(v.slice(2));
      } else if (/^\d{1,2}$/.test(s)) {
        h = Number(s);
        m = 0;
      } else {
        return String(input);
      }
      if (Number.isNaN(h) || Number.isNaN(m)) return '';
      if (mer) {
        if (mer === 'pm' && h < 12) h += 12;
        if (mer === 'am' && h === 12) h = 0;
      }
      h = Math.max(0, Math.min(23, h));
      m = Math.max(0, Math.min(59, m));
      const outMer = h >= 12 ? 'pm' : 'am';
      const h12 = (h % 12) === 0 ? 12 : (h % 12);
      const mm = String(m).padStart(2, '0');
      return `${h12}:${mm} ${outMer}`;
    };
    const parseMinutes = (input) => {
      if (!input) return null;
      let s = String(input).trim();
      if (!s) return null;
      s = s.replace(/\./g, ':').replace(/\s+/g, '');
      let mer = null;
      const lower = s.toLowerCase();
      if (/(am|pm)$/.test(lower)) { mer = lower.slice(-2); s = lower.slice(0, -2); }
      let h = 0; let m = 0;
      if (/^\d{1,2}:\d{2}$/.test(s)) { const parts = s.split(':'); h = Number(parts[0]); m = Number(parts[1]); }
      else if (/^\d{3,4}$/.test(s)) { const v = s.padStart(4, '0'); h = Number(v.slice(0,2)); m = Number(v.slice(2)); }
      else if (/^\d{1,2}$/.test(s)) { h = Number(s); m = 0; }
      else { return null; }
      if (Number.isNaN(h) || Number.isNaN(m)) return null;
      if (mer) { if (mer === 'pm' && h < 12) h += 12; if (mer === 'am' && h === 12) h = 0; }
      h = Math.max(0, Math.min(23, h)); m = Math.max(0, Math.min(59, m));
      return h * 60 + m;
    };
    const fmtFromMinutes = (mins) => {
      if (mins == null) return '';
      let v = ((mins % 1440) + 1440) % 1440;
      const h = Math.floor(v / 60); const m = v % 60;
      const outMer = h >= 12 ? 'pm' : 'am';
      const h12 = (h % 12) === 0 ? 12 : (h % 12);
      const mm = String(m).padStart(2, '0');
      return `${h12}:${mm} ${outMer}`;
    };
    const startTime = formatTime(js.event_start);
    const endTime = formatTime(js.event_end);
    const timeLine = startTime && endTime
      ? `Event time: ${startTime} – ${endTime}`
      : (startTime ? `Event time: ${startTime}` : (endTime ? `Event end: ${endTime}` : ''));
    // Call time: 1 hour 15 minutes before start
    const startMins = parseMinutes(js.event_start);
    const callTime = startMins != null ? fmtFromMinutes(startMins - 75) : '';
    const eventTime = include.event_time !== false ? String((values.event_time || '')).trim() : '';
    const callTimeLine = include.call_time !== false ? String((values.call_time || '')).trim() : '';
    const scheduleText = include.schedule ? String((values.schedule || '')).trim() : '';
    const scheduleCombined = [eventTime, callTimeLine, scheduleText].filter(Boolean).join('\n');
    if (scheduleCombined) {
      parts.push(`<div class="section"><div class="label">Schedule</div><div class="value">${safeHtml(scheduleCombined).replace(/\n/g, '<br>')}</div></div>`);
    }
  }

  // Personnel lineup
  if (include.personnel_lineup) {
    const line = String((values.personnel_lineup || '')).trim();
    if (line) parts.push(`<div class=\"section\"><div class=\"label\">Personnel</div><div class=\"value\">${safeHtml(line).replace(/\\n/g, '<br>')}</div></div>`);
  }

  // Setlist / Repertoire
  if (include.repertoire) {
    const rep = String((values.repertoire || '')).trim();
    if (rep) parts.push(`<div class=\"section\"><div class=\"label\">Setlist / Repertoire</div><div class=\"value\">${safeHtml(rep).replace(/\\n/g, '<br>')}</div></div>`);
  }

  // Dress & Kit
  if (include.dress_code) {
    const dress = values.dress_code || '';
    if (dress.trim()) parts.push(`<div class="section"><div class="label">Dress code</div><div class="value">${safeHtml(dress)}</div></div>`);
  }
  if (include.kit_notes) {
    const kit = values.kit_notes || '';
    if (kit.trim()) parts.push(`<div class="section"><div class="label">Kit</div><div class="value">${safeHtml(kit)}</div></div>`);
  }

  // Contacts
  if (include.contacts) {
    const c1 = values.contractor_name || '';
    const c1p = values.contractor_phone || '';
    const vcn = values.venue_contact_name || '';
    const vcp = values.venue_contact_phone || '';
    const lines = [];
    if (c1 || c1p) lines.push([c1, c1p].filter(Boolean).map(safeHtml).join(' · '));
    if (vcn || vcp) lines.push([vcn, vcp].filter(Boolean).map(safeHtml).join(' · '));
    if (lines.length) {
      parts.push(`<div class="section"><div class="label">Contacts</div><div class="value">${lines.join('<br>')}</div></div>`);
    }
  }

  // Notes
  if (include.notes) {
    const notes = values.notes || '';
    if (notes.trim()) parts.push(`<div class="section"><div class="label">Notes</div><div class="value">${safeHtml(notes).replace(/\n/g, '<br>')}</div></div>`);
  }

  const html = `<!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <style>
      /* Base */
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
        color: #0f172a;
        margin: 24px;
        line-height: 1.5;
        background: #ffffff;
      }
      .doc { max-width: 780px; margin: 0 auto; }
      .title { font-size: 22px; font-weight: 800; color: #0f172a; margin: 0 0 ${compact ? '12px' : '18px'} 0; }

      /* Sections */
      .section { padding: ${compact ? '8px 10px' : '12px 14px'}; border: 1px solid #e5e7eb; background-color: #f8fafc; border-radius: 10px; margin: ${compact ? '8px 0' : '12px 0'}; }
      .label { font-size: ${compact ? '11px' : '12px'}; text-transform: uppercase; letter-spacing: 0.04em; color: #475569; margin: 0 0 ${compact ? '4px' : '6px'} 0; }
      .value { font-size: ${compact ? '13px' : '14px'}; color: #0f172a; white-space: pre-wrap; }
      .footer { margin-top: 28px; font-size: 11px; color: #64748b; }

      /* Let content flow across page breaks naturally */
      @media print {
        .section { break-inside: auto; page-break-inside: auto; }
        .value { overflow-wrap: anywhere; word-break: break-word; }
      }
    </style>
  </head>
  <body>
    <div class="doc">
      <div class="title">${safeHtml(header)}</div>
      ${parts.join('\n')}
    </div>
  </body>
  </html>`;

  // Determine save path
  const ensured = await module.exports.ensureJobsheetFolder({ businessId, jobsheetId });
  const folderPath = ensured?.folder_path || ensured?.path || '';
  if (!folderPath) throw new Error('Unable to resolve jobsheet folder');
  // Name PDF with human date suffix when available (e.g., "Gig Info - 08 Oct 2025.pdf")
  const rawDate = js.event_date ? String(js.event_date).trim() : '';
  const datePart = rawDate ? formatGigDate(rawDate).replace(/,/g, '') : '';
  const base = datePart ? `Gig Info - ${datePart}` : 'Gig Info';
  let target = path.join(folderPath, `${base}.pdf`);
  let n = 2;
  while (await pathExists(target)) {
    target = path.join(folderPath, `${base} (${n}).pdf`);
    n += 1;
    if (n > 999) break;
  }
  return { html, targetPath: target };
}

async function syncJobsheetOutputs(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required to sync outputs.');
  }

  const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
  const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;

  const hintPaths = Array.isArray(options.hintPaths) ? options.hintPaths.filter(Boolean) : [];
  const explicitDirectories = Array.isArray(options.directories) ? options.directories.filter(Boolean) : [];
  const snapshot = options.jobsheetSnapshot && typeof options.jobsheetSnapshot === 'object'
    ? { ...options.jobsheetSnapshot }
    : {};

  const extensionsInput = Array.isArray(options.extensions) && options.extensions.length
    ? options.extensions
    : ['.pdf'];

  const normalizedExtensions = extensionsInput
    .map(ext => {
      if (!ext) return null;
      const trimmed = ext.toString().trim();
      if (!trimmed) return null;
      return trimmed.startsWith('.') ? trimmed.toLowerCase() : `.${trimmed.toLowerCase()}`;
    })
    .filter(Boolean);

  if (!normalizedExtensions.length) {
    return { added: 0, records: [] };
  }

  const business = await db.getBusinessById(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured for this business.');
  }

  const rootPath = path.resolve(business.save_path);
  const directories = new Set();

  explicitDirectories.forEach(dir => {
    try {
      if (dir) directories.add(path.resolve(dir));
    } catch (_err) {
      // ignore invalid paths
    }
  });

  hintPaths.forEach(filePath => {
    if (!filePath || typeof filePath !== 'string') return;
    try {
      const absolute = path.resolve(filePath);
      directories.add(path.dirname(absolute));
    } catch (_err) {
      // ignore resolve errors
    }
  });

  if (!directories.size) {
    try {
      const payload = {
        business_id: businessId,
        jobsheet_id: jobsheetId,
        jobsheet_snapshot: snapshot,
        client_override: {},
        event_override: {},
        pricing_snapshot: {}
      };
      const context = buildContext(payload, business);
      const naming = buildFileName(context, payload, { label: 'Workbook' });
      const expectedDirectory = buildOutputDirectory(business, context, payload, naming.folderName);
      directories.add(expectedDirectory);
    } catch (err) {
      console.warn('Unable to determine expected jobsheet directory', err);
    }
  }

  if (!directories.size) {
    return { added: 0, records: [] };
  }

  const results = [];
  let added = 0;

  for (const dir of directories) {
    if (!dir) continue;
    let resolvedDir;
    try {
      resolvedDir = path.resolve(dir);
    } catch (_err) {
      continue;
    }

    if (!isSubPath(rootPath, resolvedDir)) {
      continue;
    }

    if (!(await pathExists(resolvedDir))) {
      continue;
    }

    let entries;
    try {
      entries = await fs.promises.readdir(resolvedDir, { withFileTypes: true });
    } catch (err) {
      console.warn('Unable to read output directory', resolvedDir, err);
      continue;
    }

    for (const entry of entries) {
      if (!entry.isFile()) continue;
      const ext = path.extname(entry.name).toLowerCase();
      if (!normalizedExtensions.includes(ext)) continue;

      const absolutePath = path.join(resolvedDir, entry.name);
      try {
        const existing = await db.getDocumentByFilePath(businessId, absolutePath);
        if (existing) {
          continue;
        }

        const stats = await fs.promises.stat(absolutePath);
        const documentDate = stats?.mtime instanceof Date && !Number.isNaN(stats.mtime.valueOf())
          ? stats.mtime.toISOString()
          : new Date().toISOString();

        const inserted = await db.addDocument({
          business_id: businessId,
          jobsheet_id: jobsheetId,
          doc_type: 'pdf_export',
          status: 'exported',
          total_amount: null,
          balance_due: null,
          due_date: null,
          file_path: absolutePath,
          client_name: snapshot.client_name || snapshot.client || null,
          event_name: snapshot.event_type || snapshot.event_name || null,
          event_date: snapshot.event_date || null,
          document_date: documentDate,
          definition_key: null,
          invoice_variant: null
        });

        added += 1;
        results.push({ document_id: inserted?.id || null, file_path: absolutePath });
      } catch (err) {
        console.error('Failed to record exported document', absolutePath, err);
      }
    }
  }

  return { added, records: results };
}


async function deleteDocument(documentId, options = {}) {
  const id = Number(documentId);
  if (!Number.isInteger(id)) {
    throw new Error('A valid document id is required.');
  }

  const removeFile = options.removeFile === true;
  let record = null;
  try {
    record = await db.getDocumentById(id);
  } catch (err) {
    // ignore lookup errors and continue with deletion
  }

  if (record && record.is_locked) {
    throw new Error('This document is locked and cannot be modified.');
  }

  if (removeFile && record?.file_path) {
    try {
      await fs.promises.unlink(record.file_path);
    } catch (err) {
      if (err && err.code !== 'ENOENT') {
        throw new Error(`Unable to delete file: ${err.message || err}`);
      }
    }
  }

  // Capture potential rollback for invoice numbering
  const wasInvoice = (record?.doc_type || '').toLowerCase() === 'invoice';
  const businessId = record?.business_id != null ? Number(record.business_id) : null;
  const deletedNumber = record?.number != null ? Number(record.number) : null;

  await db.deleteDocument(id);

  // If deleting an invoice, ensure last_invoice_number remains consistent (set to current max)
  if (wasInvoice && Number.isInteger(businessId)) {
    try {
      const maxNum = await db.getMaxInvoiceNumber(businessId);
      const last = Number.isInteger(Number(maxNum)) ? Number(maxNum) : 0;
      await db.setLastInvoiceNumber(businessId, last);
    } catch (_err) {}
  }

  return { ok: true };
}

async function preflightPdfExport(options = {}) {
  const providedPath = options.filePath || options.file_path;
  if (!providedPath || typeof providedPath !== 'string' || !providedPath.trim()) {
    throw new Error('filePath is required for preflight.');
  }

  const normalizedPath = path.resolve(providedPath.trim());
  const dir = path.dirname(normalizedPath);
  const baseName = path.basename(normalizedPath, path.extname(normalizedPath));
  const targetPdfPath = path.join(dir, `${baseName}.pdf`);

  const checks = [];
  const addCheck = (name, ok, message) => checks.push({ name, ok: Boolean(ok), message: message || '' });

  // Check workbook file exists and is readable
  try {
    await fs.promises.access(normalizedPath, fs.constants.R_OK);
    addCheck('workbook_exists', true);
  } catch (err) {
    addCheck('workbook_exists', false, `Workbook not accessible: ${err?.message || err}`);
  }

  // Check destination directory is writable
  try {
    await fs.promises.access(dir, fs.constants.W_OK);
    // Try a temp write/delete to be certain (especially on cloud folders)
    const tmp = path.join(dir, `.preflight_${Date.now()}_${Math.random().toString(36).slice(2)}.tmp`);
    await fs.promises.writeFile(tmp, 'ok');
    await fs.promises.unlink(tmp);
    addCheck('destination_writable', true);
  } catch (err) {
    addCheck('destination_writable', false, `Destination not writable: ${err?.message || err}`);
  }

  // Check Excel Apple Events permission / availability
  try {
    await new Promise((resolve, reject) => {
      execFile('osascript', ['-e', 'tell application "Microsoft Excel" to version'], { timeout: 15000 }, (error, stdout, stderr) => {
        if (error) {
          const message = (stderr || stdout || error.message || '').toString().trim();
          reject(new Error(message || 'Unable to communicate with Microsoft Excel'));
          return;
        }
        resolve();
      });
    });
    addCheck('excel_automation', true);
  } catch (err) {
    addCheck('excel_automation', false, `Excel automation failed: ${err?.message || err}`);
  }

  // Attempt to open and close the workbook (no save), to catch Excel-specific open errors
  try {
    const osaArgs = [
      '-e', 'on run argv',
      '-e', 'if (count of argv) < 1 then error "Missing path"',
      '-e', 'set workbookPosixPath to item 1 of argv',
      '-e', 'set workbookHfs to (POSIX file workbookPosixPath) as text',
      '-e', 'tell application "Microsoft Excel"',
      '-e', 'launch',
      // Keep Excel in background: do not activate, hide UI and alerts
      '-e', 'try',
      '-e', 'set visible to false',
      '-e', 'set display alerts to false',
      '-e', 'end try',
      '-e', 'set wb to missing value',
      '-e', 'try',
      '-e', 'set wb to open workbook workbook file name workbookHfs',
      '-e', 'end try',
      '-e', 'repeat with i from 1 to 50',
      '-e', 'if wb is not missing value then exit repeat',
      '-e', 'delay 0.1',
      '-e', 'try',
      '-e', 'set wb to active workbook',
      '-e', 'end try',
      '-e', 'end repeat',
      '-e', 'if wb is missing value then error "Unable to open workbook"',
      '-e', 'try',
      '-e', 'close workbook wb saving no',
      '-e', 'end try',
      '-e', 'end tell',
      // Ensure Excel is not frontmost as a fallback
      '-e', 'try',
      '-e', 'tell application "System Events" to set frontmost of process "Microsoft Excel" to false',
      '-e', 'end try',
      '-e', 'end run',
      normalizedPath
    ];
    await new Promise((resolve, reject) => {
      execFile('osascript', osaArgs, { timeout: 45000 }, (error, stdout, stderr) => {
        if (error) {
          const message = (stderr || stdout || error.message || '').toString().trim();
          reject(new Error(message || 'Excel failed to open the workbook'));
          return;
        }
        resolve();
      });
    });
    addCheck('workbook_openable', true);
  } catch (err) {
    addCheck('workbook_openable', false, `Unable to open in Excel: ${err?.message || err}`);
  }

  const ok = checks.every(c => c.ok);
  const firstFailure = checks.find(c => !c.ok);
  const message = ok ? 'Preflight OK' : (firstFailure?.message || 'Preflight failed');

  return {
    ok,
    workbook_path: normalizedPath,
    target_pdf_path: targetPdfPath,
    checks,
    message
  };
}

// Compose an email in Apple Mail with attachments. Leaves draft open for user to send.
async function composeMailDraft(options = {}) {
  const to = (options.to || '').toString();
  const subject = (options.subject || '').toString();
  const body = (options.body || '').toString();
  const fromEmail = (options.fromEmail || '').toString();
  // Signature handling intentionally omitted — rely on Mail's default per-account signature
  const attachments = Array.isArray(options.attachments) ? options.attachments.filter(Boolean) : [];

  // Compose using UI "New Message" to preserve default signature
  const argsUi = [
    '-e', 'on run argv',
    '-e', 'set theTo to item 1 of argv',
    '-e', 'set theSubject to item 2 of argv',
    '-e', 'set theBody to item 3 of argv',
    '-e', 'set theCount to item 4 of argv as integer',
    '-e', 'set theAttachments to {}',
    '-e', 'try',
    '-e', '  repeat with i from 1 to theCount',
    '-e', '    set p to item (4 + i) of argv',
    '-e', '    set end of theAttachments to p',
    '-e', '  end repeat',
    '-e', 'end try',
    '-e', 'tell application "Mail" to activate',
    '-e', 'tell application "System Events" to tell process "Mail" to click menu item "New Message" of menu "File" of menu bar 1',
    '-e', 'delay 0.3',
    '-e', 'tell application "Mail"',
    '-e', '  set msg to front message',
    '-e', '  if (theTo is not "") then tell msg to make new to recipient with properties {address:theTo}',
    '-e', '  if (theSubject is not "") then set subject of msg to theSubject',
    // Do not touch content or sender — preserve default signature
    '-e', '  try',
    '-e', '    repeat with p in theAttachments',
    '-e', '      set f to (POSIX file (contents of p)) as alias',
    '-e', '      tell msg to make new attachment with properties {file name:f} at after the last paragraph',
    '-e', '    end repeat',
    '-e', '  end try',
    '-e', '  activate',
    '-e', 'end tell',
    '-e', 'end run'
  ];

  // Fallback compose without UI (signature may be None)
  const args = [
    '-e', 'on run argv',
    '-e', 'set theTo to item 1 of argv',
    '-e', 'set theSubject to item 2 of argv',
    '-e', 'set theBody to item 3 of argv',
    '-e', 'set theFrom to item 4 of argv',
    '-e', 'set theCount to item 5 of argv as integer',
    '-e', 'set theAttachments to {}',
    '-e', 'try',
    '-e', '  repeat with i from 1 to theCount',
    '-e', '    set p to item (5 + i) of argv',
    '-e', '    set end of theAttachments to p',
    '-e', '  end repeat',
    '-e', 'end try',
    '-e', 'tell application "Mail"',
    '-e', '  activate',
    '-e', '  set msg to make new outgoing message with properties {visible:true, subject:theSubject, content:theBody & return & return}',
    '-e', '  if (theTo is not "") then tell msg to make new to recipient with properties {address:theTo}',
    '-e', '  if (theFrom is not "") then tell msg to set sender to theFrom',
    '-e', '  try',
    '-e', '    repeat with p in theAttachments',
    '-e', '      set f to (POSIX file (contents of p)) as alias',
    '-e', '      tell msg to make new attachment with properties {file name:f} at after the last paragraph',
    '-e', '    end repeat',
    '-e', '  end try',
    '-e', '  activate',
    '-e', 'end tell',
    '-e', 'end run'
  ];

  const payload = [to, subject, body, String(attachments.length)].concat(attachments);
  // Try UI method first
  try {
    await new Promise((resolve, reject) => {
      execFile('osascript', argsUi.concat(payload), { timeout: 30000 }, (error, stdout, stderr) => {
        if (error) {
          const msg = (stderr || stdout || error.message || '').toString();
          return reject(new Error(msg.trim() || 'UI compose failed'));
        }
        resolve();
      });
    });
  } catch (_err) {
    // Do not fallback to programmatic compose, as it can clear the signature
    throw _err;
  }
  return { ok: true };
}

// Alternative compose that relies on mailto: (preserves default signature) and then attaches files
async function composeMailDraft_mailto(options = {}) {
  const to = (options.to || '').toString();
  const subject = (options.subject || '').toString();
  const body = (options.body || '').toString();
  const attachments = Array.isArray(options.attachments) ? options.attachments.filter(Boolean) : [];

  const params = [];
  if (subject) params.push(`subject=${encodeURIComponent(subject)}`);
  if (body) params.push(`body=${encodeURIComponent(body)}`);
  const mailtoUrl = `mailto:${encodeURIComponent(to)}${params.length ? `?${params.join('&')}` : ''}`;

  await new Promise((resolve, reject) => {
    execFile('open', [mailtoUrl], { timeout: 15000 }, (error, stdout, stderr) => {
      if (error) {
        const msg = (stderr || stdout || error.message || '').toString();
        reject(new Error(msg.trim() || 'Unable to open Mail compose'));
        return;
      }
      resolve();
    });
  });

  let usedClipboard = false;
  if (attachments.length) {
    const osaAttach = [
      '-e', 'on run argv',
      '-e', 'set theSubject to item 1 of argv',
      '-e', 'set theCount to item 2 of argv as integer',
      '-e', 'set theAttachments to {}',
      '-e', 'repeat with i from 1 to theCount',
      '-e', '  set end of theAttachments to item (2 + i) of argv',
      '-e', 'end repeat',
      '-e', 'tell application "Mail"',
      '-e', '  set newMsg to missing value',
      '-e', '  repeat 50 times',
      '-e', '    try',
      '-e', '      set msgs to every outgoing message',
      '-e', '      if (count of msgs) > 0 then',
      '-e', '        if (theSubject is not "") then',
      '-e', '          set matches to {}',
      '-e', '          repeat with m in msgs',
      '-e', '            try',
      '-e', '              if (subject of m as string) contains theSubject then set end of matches to m',
      '-e', '            end try',
      '-e', '          end repeat',
      '-e', '          if (count of matches) > 0 then set newMsg to item 1 of matches',
      '-e', '        end if',
      '-e', '        if newMsg is missing value then set newMsg to item 1 of msgs',
      '-e', '      end if',
      '-e', '    end try',
      '-e', '    if newMsg is not missing value then exit repeat',
      '-e', '    delay 0.2',
      '-e', '  end repeat',
      '-e', '  if newMsg is missing value then error "No outgoing message found"',
      '-e', '  try',
      '-e', '    repeat with p in theAttachments',
      '-e', '      set f to (POSIX file (contents of p)) as alias',
      '-e', '      tell newMsg to make new attachment with properties {file name:f} at after the last paragraph',
      '-e', '    end repeat',
      '-e', '  end try',
      '-e', '  activate',
      '-e', 'end tell',
      '-e', 'end run'
    ];
    const payload = [subject, String(attachments.length)].concat(attachments);
    try {
      await new Promise((resolve, reject) => {
        execFile('osascript', osaAttach.concat(payload), { timeout: 30000 }, (error, stdout, stderr) => {
          if (error) {
            const msg = (stderr || stdout || error.message || '').toString();
            return reject(new Error(msg.trim() || 'Unable to attach files'));
          }
          resolve();
        });
      });
    } catch (_attachErr) {
      // Fallback: copy files to clipboard so user can paste into the draft manually (Cmd+V)
      try {
        const osaClipboard = [
          '-e', 'on run argv',
          '-e', 'set theCount to item 1 of argv as integer',
          '-e', 'set fileList to {}',
          '-e', 'repeat with i from 1 to theCount',
          '-e', '  set p to item (1 + i) of argv',
          '-e', '  set end of fileList to ((POSIX file p) as alias)',
          '-e', 'end repeat',
          '-e', 'set the clipboard to fileList',
          '-e', 'end run'
        ];
        const clipPayload = [String(attachments.length)].concat(attachments);
        await new Promise((resolve) => {
          execFile('osascript', osaClipboard.concat(clipPayload), { timeout: 10000 }, () => resolve());
        });
      } catch (_) {
        // ignore clipboard errors
      }
      usedClipboard = true; // files copied; user can paste into Mail
      // Continue without throwing to keep the compose workflow usable
    }
  }

  return { ok: true, used_clipboard: usedClipboard };
}

// Locate already-exported PDFs needed for the Booking Pack within the job folder
async function getBookingPackPdfs(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
  const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;
  const snapshot = options.jobsheetSnapshot && typeof options.jobsheetSnapshot === 'object' ? { ...options.jobsheetSnapshot } : {};

  if (!Number.isInteger(businessId)) throw new Error('businessId is required');

  // Resolve job folder
  const business = await db.getBusinessById(businessId);
  if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');
  const payload = { business_id: businessId, jobsheet_id: jobsheetId, jobsheet_snapshot: snapshot };
  const context = buildContext(payload, business);
  const folderPath = buildOutputDirectory(business, context, payload, 'Documents');

  // Build quick index of PDFs in the folder
  let entries = [];
  try { entries = await fs.promises.readdir(folderPath, { withFileTypes: true }); } catch (_) { entries = []; }
  const pdfs = [];
  for (const e of entries) {
    if (!e || !e.name || e.isDirectory()) continue;
    if (!e.name.toLowerCase().endsWith('.pdf')) continue;
    pdfs.push({ name: e.name, path: path.join(folderPath, e.name) });
  }

  const findFirst = (regexes) => {
    for (const r of regexes) {
      const hit = pdfs.find(p => r.test(p.name));
      if (hit) return hit.path;
    }
    return '';
  };

  // Heuristics
  const schedulePdf = findFirst([/schedule/i, /booking\s*schedule/i]);
  const termsPdf = findFirst([/t\s*&\s*c/i, /tandc/i, /tnc/i, /terms/i, /terms\s*&\s*conditions/i, /conditions/i]);
  // Prefer deposit-specific filenames first, then generic invoice pattern (excluding balance)
  let depositPdf = findFirst([
    /deposit/i,
    /\bdep\b/i
  ]);
  if (!depositPdf) {
    // fallback to generic invoice file if it doesn't look like a balance
    const generic = findFirst([/\(\s*INV[-\s]?\d+\s*\)\.pdf$/i]);
    if (generic && !/balance/i.test(path.basename(generic))) depositPdf = generic;
  }

  // As a stronger signal, consult DB for an invoice PDF for this jobsheet (if available)
  try {
    const docs = await db.getDocuments({ businessId });
    const byJob = Array.isArray(docs) ? docs.filter(d => (d?.jobsheet_id != null ? Number(d.jobsheet_id) : null) === jobsheetId) : [];
    const inv = byJob.find(d => (d?.doc_type || '').toLowerCase() === 'invoice' && (d?.invoice_variant || '').toLowerCase() === 'deposit' && d?.file_path && d.file_path.toLowerCase().endsWith('.pdf'));
    if (inv && inv.file_path) depositPdf = inv.file_path;
  } catch (_) {}

  return {
    ok: true,
    folder_path: folderPath,
    schedule_pdf: schedulePdf || '',
    terms_pdf: termsPdf || '',
    deposit_pdf: depositPdf || ''
  };
}

// --- Graph helpers for in-app sending ---
const GRAPH_SCOPES = ['https://graph.microsoft.com/Mail.Send', 'offline_access', 'openid', 'profile', 'User.Read'];
function getGraphAuthority() {
  const tenant = settings.graph_tenant_id;
  return `https://login.microsoftonline.com/${tenant}`;
}
function getClientId() {
  return settings.graph_client_id;
}
function getCachePath() {
  try { return path.join(os.homedir(), '.invoice_master_msal_cache.json'); } catch (_) { return path.join(__dirname, '.msal_cache.json'); }
}

function createPublicClient() {
  const pca = new msal.PublicClientApplication({ auth: { clientId: getClientId(), authority: getGraphAuthority() } });
  try {
    const raw = fs.readFileSync(getCachePath(), 'utf-8');
    pca.getTokenCache().deserialize(raw);
  } catch (_) {}
  return pca;
}
function persistCache(pca) {
  try {
    const raw = pca.getTokenCache().serialize();
    fs.writeFileSync(getCachePath(), raw, 'utf-8');
  } catch (_) {}
}
async function acquireGraphToken() {
  const pca = createPublicClient();
  const cache = pca.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts && accounts.length) {
    try {
      const resp = await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: accounts[0] });
      persistCache(pca);
      return resp.accessToken;
    } catch (_) {}
  }
  const resp = await pca.acquireTokenByDeviceCode({
    scopes: GRAPH_SCOPES,
    deviceCodeCallback: (info) => {
      try {
        const uri = info.verificationUri || 'https://microsoft.com/devicelogin';
        // Open the verification URL in default browser
        execFile('open', [uri], () => {});
        // Copy the user code to the clipboard for easy pasting
        const osaClipboard = [
          '-e', `set the clipboard to \"${info.userCode}\"`
        ];
        execFile('osascript', osaClipboard, () => {});
        // Show a simple dialog with the code
        const message = `A Microsoft sign-in is required.\n\n1) A browser window has opened.\n2) When prompted, paste this code: ${info.userCode}`;
        execFile('osascript', ['-e', `display dialog ${JSON.stringify(message)} buttons {"OK"} giving up after 15`], () => {});
      } catch (_e) {
        // Fall back to console log
        console.log(info.message);
      }
    }
  });
  persistCache(pca);
  return resp.accessToken;
}
async function graphSendMailRaw({ to, subject, body, attachments }) {
  const token = await acquireGraphToken();
  const atts = [];
  for (const p of (attachments || [])) {
    try {
      const name = path.basename(p);
      const data = fs.readFileSync(p);
      atts.push({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: data.toString('base64') });
    } catch (_) {}
  }
  const payload = {
    message: {
      subject,
      body: { contentType: 'Text', content: body || '' },
      toRecipients: [{ emailAddress: { address: to } }],
      attachments: atts
    },
    saveToSentItems: true
  };
  const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Graph send failed: ${res.status} ${t}`);
  }
}
async function sendBookingPackViaGraph(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');
  const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
  const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;
  const snapshot = options.jobsheetSnapshot && typeof options.jobsheetSnapshot === 'object' ? { ...options.jobsheetSnapshot } : {};

  const assets = await getBookingPackPdfs({ businessId, jobsheetId, jobsheetSnapshot: snapshot });
  const files = [assets.schedule_pdf, assets.terms_pdf, assets.deposit_pdf].filter(Boolean);
  if (!files.length) throw new Error('No booking pack PDFs found in the job folder.');
  const to = (snapshot.client_email || '').trim();
  if (!to) throw new Error('Client email missing on jobsheet');
  const subject = `Booking pack – ${(snapshot.client_name || 'Client')} – ${formatDisplayDate(snapshot.event_date)}`;
  const firstName = (snapshot.client_name || '').trim().split(/\s+/)[0] || 'there';
  const body = `Hi ${firstName},\n\nAttached are your booking schedule, T&Cs, and deposit invoice. The deposit is payable on contract signing.\n\nThanks,\nMotti`;
  await graphSendMailRaw({ to, subject, body, attachments: files });
  try {
    await db.logEmail({ business_id: businessId, jobsheet_id: jobsheetId, to, cc: '', bcc: '', subject, body, attachments: files, provider: 'graph', status: 'sent', message_id: null });
  } catch (_) {}
  return { ok: true };
}

async function sendMailViaGraph(options = {}) {
  const to = (options.to || '').toString().trim();
  if (!to) throw new Error('Recipient (to) is required');
  const subject = (options.subject || '').toString();
  const body = (options.body || '').toString();
  const isHtml = options.is_html === true;
  const skipLog = options.skipLog === true;
  const cc = normalizeRecipientList(options.cc);
  const bcc = normalizeRecipientList(options.bcc);
  const attachments = Array.isArray(options.attachments) ? options.attachments.filter(Boolean) : [];

  const token = await acquireGraphToken();
  const atts = [];
  for (const p of attachments) {
    try {
      const name = path.basename(p);
      const data = fs.readFileSync(p);
      atts.push({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: data.toString('base64') });
    } catch (_) {}
  }
  const toList = to.split(/[,;]+/).map(s => s.trim()).filter(Boolean).map(address => ({ emailAddress: { address } }));
  const ccList = cc.map(address => ({ emailAddress: { address } }));
  const bccList = bcc.map(address => ({ emailAddress: { address } }));

  const payload = {
    message: {
      subject,
      body: { contentType: isHtml ? 'HTML' : 'Text', content: body || '' },
      toRecipients: toList,
      ccRecipients: ccList,
      bccRecipients: bccList,
      attachments: atts
    },
    saveToSentItems: true
  };
  const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Graph send failed: ${res.status} ${t}`);
  }
  if (!skipLog) {
    try {
      await db.logEmail({
        business_id: options.business_id ?? null,
        jobsheet_id: options.jobsheet_id ?? null,
        to,
        cc,
        bcc,
        subject,
        body,
        attachments,
        provider: 'graph',
        status: 'sent',
        message_id: null
      });
      broadcastJobsheetChange({
        type: 'email-log-updated',
        businessId: options.business_id != null ? Number(options.business_id) : null,
        jobsheetId: options.jobsheet_id != null ? Number(options.jobsheet_id) : null
      });
    } catch (_) {}
  }
  return { ok: true };
}

async function scheduleMailViaGraph(options = {}) {
  const to = (options.to || '').toString().trim();
  if (!to) throw new Error('Recipient (to) is required');

  const subject = (options.subject || '').toString();
  const body = (options.body || '').toString();
  const isHtml = options.is_html !== false; // default to HTML since composer sends HTML
  const cc = normalizeRecipientList(options.cc);
  const bcc = normalizeRecipientList(options.bcc);
  const attachments = Array.isArray(options.attachments) ? options.attachments.filter(Boolean) : [];

  const sendAtInput = options.send_at ?? options.schedule_at ?? options.sendAt;
  if (!sendAtInput) throw new Error('send_at is required for scheduling');

  const sendAtDate = sendAtInput instanceof Date ? sendAtInput : new Date(sendAtInput);
  if (Number.isNaN(sendAtDate.valueOf())) throw new Error('Invalid send_at value');

  const now = Date.now();
  if (sendAtDate.getTime() < now + 30 * 1000) {
    throw new Error('Scheduled send time must be at least 30 seconds in the future');
  }

  const businessId = options.business_id ?? options.businessId ?? null;
  const jobsheetId = options.jobsheet_id ?? options.jobsheetId ?? null;

  const emailLogId = await db.logEmail({
    business_id: businessId,
    jobsheet_id: jobsheetId,
    to,
    cc,
    bcc,
    subject,
    body,
    attachments,
    provider: 'graph',
    status: 'scheduled',
    message_id: null,
    sent_at: sendAtDate
  });

  const scheduledId = await db.queueScheduledEmail({
    email_log_id: emailLogId,
    business_id: businessId,
    jobsheet_id: jobsheetId,
    to,
    cc,
    bcc,
    subject,
    body,
    attachments,
    is_html: isHtml,
    send_at: sendAtDate
  });

  broadcastJobsheetChange({
    type: 'email-log-updated',
    businessId: businessId != null ? Number(businessId) : null,
    jobsheetId: jobsheetId != null ? Number(jobsheetId) : null
  });

  ensureScheduledMailWorker();

  return { ok: true, scheduled_email_id: scheduledId, email_log_id: emailLogId };
}

async function resolveTemplateDefaultAttachments(options = {}) {
  const templateKeyRaw = options.templateKey ?? options.template_key;
  const templateKey = templateKeyRaw ? String(templateKeyRaw).trim().toLowerCase() : '';
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
  const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;

  if (!templateKey) return { attachments: [] };
  if (!Number.isInteger(businessId)) return { attachments: [] };

  const attachments = [];

  let jobFolderFiles = [];
  try {
    jobFolderFiles = await module.exports.listJobFolderFiles({ businessId, jobsheetId, extensionPattern: '\\.(pdf)$' });
  } catch (_) {
    jobFolderFiles = [];
  }

  const findInFolder = (patterns) => {
    for (const pattern of patterns) {
      const match = jobFolderFiles.find(file => pattern.test(file?.name || ''));
      if (match?.path) return match.path;
    }
    return '';
  };

  const resolveFromDocuments = async (predicate) => {
    try {
      const response = await listJobsheetDocuments({ businessId, jobsheetId });
      const docs = Array.isArray(response?.documents) ? response.documents : [];
      const filtered = jobsheetId != null
        ? docs.filter(doc => (doc?.jobsheet_id != null ? Number(doc.jobsheet_id) : null) === jobsheetId)
        : docs;
      const match = filtered.find(predicate);
      if (match?.file_path) return match.file_path;
      return '';
    } catch (_err) {
      return '';
    }
  };

  switch (templateKey) {
    case 'booking_pack': {
      try {
        const bundle = await getBookingPackPdfs({ businessId, jobsheetId });
        attachments.push(bundle?.schedule_pdf || '', bundle?.terms_pdf || '', bundle?.deposit_pdf || '');
      } catch (_) {}
      if (!attachments.filter(Boolean).length) {
        const schedule = findInFolder([
          /schedule/i,
          /booking\s*schedule/i,
          /itinerary/i
        ]);
        const terms = findInFolder([
          /t\s*&\s*c/i,
          /terms/i,
          /conditions/i,
          /tandc/i,
          /tnc/i
        ]);
        const deposit = findInFolder([
          /deposit/i,
          /\bdep\b/i,
          /invoice[-_\s]*deposit/i
        ]);
        attachments.push(schedule || '', terms || '', deposit || '');
      }
      break;
    }
    case 'invoice_balance': {
      let pathMatch = findInFolder([
        /balance/i,
        /invoice[_\s-]*balance/i,
        /bal[-_\s]*inv/i
      ]);
      if (!pathMatch) {
        pathMatch = await resolveFromDocuments(doc => {
          const defKey = String(doc?.definition_key || '').toLowerCase();
          const docType = String(doc?.doc_type || '').toLowerCase();
          const variant = String(doc?.invoice_variant || '').toLowerCase();
          if (defKey === 'invoice_balance') return true;
          return docType === 'invoice' && variant === 'balance';
        });
      }
      if (pathMatch) attachments.push(pathMatch);
      break;
    }
    case 'quote': {
      let pathMatch = findInFolder([
        /quote/i,
        /quotation/i
      ]);
      if (!pathMatch) {
        pathMatch = await resolveFromDocuments(doc => {
          const defKey = String(doc?.definition_key || '').toLowerCase();
          const docType = String(doc?.doc_type || '').toLowerCase();
          if (defKey === 'quote') return true;
          return docType === 'quote';
        });
      }
      if (pathMatch) attachments.push(pathMatch);
      break;
    }
    default:
      break;
  }

  const unique = [];
  const seen = new Set();
  attachments.filter(Boolean).forEach(p => {
    const normalized = path.resolve(p);
    if (seen.has(normalized)) return;
    seen.add(normalized);
    unique.push(p);
  });

  return { attachments: unique };
}

ensureScheduledMailWorker();
module.exports = {
  normalizeTemplate,
  createDocument,
  buildMCMSDocumentHtml,
  getHtmlTemplate: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const docType = (options.docType || options.doc_type || '').toString().toLowerCase();
    if (!Number.isInteger(businessId) || !docType) return { html: '' };
    const s = readSettings();
    const map = s.html_templates_by_business || {};
    const byBiz = map[String(businessId)] || {};
    return { html: byBiz[docType] || '' };
  },
  saveHtmlTemplate: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const docType = (options.docType || options.doc_type || '').toString().toLowerCase();
    const html = (options.html || '').toString();
    if (!Number.isInteger(businessId) || !docType) return { ok: false };
    const s = readSettings();
    const map = s.html_templates_by_business || {};
    const byBiz = map[String(businessId)] || {};
    byBiz[docType] = html;
    map[String(businessId)] = byBiz;
    s.html_templates_by_business = map;
    writeSettings(s);
    return { ok: true };
  },
  // Read a lightweight snapshot of an Excel workbook for in-app editing
  readExcelSnapshot: async (options = {}) => {
    const filePath = options.filePath || options.file_path;
    const requestedSheet = options.sheetName || options.sheet || '';
    const maxRows = Number.isInteger(options.maxRows) ? options.maxRows : 100;
    const maxCols = Number.isInteger(options.maxCols) ? options.maxCols : 26; // A-Z
    if (!filePath || typeof filePath !== 'string') throw new Error('filePath is required');
    const resolved = path.resolve(filePath);
    await ensureFileAccessible(resolved);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(resolved);
    const sheetNames = wb.worksheets.map(ws => ws && ws.name).filter(Boolean);
    if (!sheetNames.length) return { sheets: [], sheet: '', cells: [] };
    const sheet = requestedSheet && sheetNames.includes(requestedSheet) ? requestedSheet : sheetNames[0];
    const ws = wb.getWorksheet(sheet);
    const rows = [];
    for (let r = 1; r <= maxRows; r++) {
      const row = [];
      for (let c = 1; c <= maxCols; c++) {
        try {
          const cell = ws.getCell(r, c);
          let val = cell && cell.value != null ? cell.value : '';
          if (typeof val === 'object' && val && 'text' in val) val = val.text;
          if (typeof val === 'object' && val && 'result' in val) val = val.result;
          row.push(val == null ? '' : val);
        } catch (_) {
          row.push('');
        }
      }
      rows.push(row);
    }
    return { sheets: sheetNames, sheet, cells: rows };
  },
  // Persist edited cells back to the Excel file
  writeExcelCells: async (options = {}) => {
    const filePath = options.filePath || options.file_path;
    const sheetName = options.sheetName || options.sheet || '';
    const changes = Array.isArray(options.changes) ? options.changes : [];
    if (!filePath || typeof filePath !== 'string') throw new Error('filePath is required');
    if (!sheetName) throw new Error('sheetName is required');
    const resolved = path.resolve(filePath);
    await ensureFileAccessible(resolved);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(resolved);
    const ws = wb.getWorksheet(sheetName);
    if (!ws) throw new Error('Sheet not found');
    for (const ch of changes) {
      const r = Number(ch.row), c = Number(ch.col);
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 1 || c < 1) continue;
      const v = ch.value;
      ws.getCell(r, c).value = v;
    }
    await wb.xlsx.writeFile(resolved);
    return { ok: true };
  },
  composeMailDraft,
  composeMailDraft_mailto,
  // MCMS: Create a numbered invoice/quote from a single Excel template without jobsheets
  // Options: { business_id, doc_type: 'invoice'|'quote', definition_key, client_override, event_override, total_amount, due_date, document_date }
  createNumberedDocument: async (options = {}) => {
    const businessId = Number(options.business_id ?? options.businessId);
    if (!Number.isInteger(businessId)) throw new Error('business_id is required');
    const rawType = String(options.doc_type || options.type || '').toLowerCase();
    if (!rawType || (rawType !== 'invoice' && rawType !== 'quote')) throw new Error('doc_type must be invoice or quote');

    const definitionKey = options.definition_key || (rawType === 'invoice' ? 'invoice_balance' : 'quote');
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');

    // Resolve template path from document definition
    const def = await db.getDocumentDefinition(businessId, definitionKey);
    const templatePath = def?.template_path || options.template_path || options.templatePath;
    if (!templatePath) throw new Error('Template path is not configured for this document. Set it in Templates.');
    const resolvedTemplate = path.resolve(templatePath);
    await ensureFileAccessible(resolvedTemplate);

    // Build context (client/event overrides supported)
    const payload = {
      business_id: businessId,
      client_override: options.client_override || {},
      event_override: options.event_override || {},
      total_amount: options.total_amount ?? null,
      balance_amount: options.total_amount ?? null,
      balance_due: options.total_amount ?? null,
      due_date: options.due_date || null,
      document_date: options.document_date || new Date().toISOString(),
      definition_key: definitionKey
    };
    const context = buildContext(payload, business);

    // Determine target folder/name
    const naming = buildFileName(context, payload, { label: rawType === 'invoice' ? 'Invoice' : 'Quote', key: definitionKey });
    const directory = buildOutputDirectory(business, context, payload, naming.folderName);
    await ensureDirectoryExists(directory);

    // Compose workbook path and write filled workbook
    const workbookPath = path.join(directory, `${naming.fileName}`);
    // Avoid overwriting existing workbook
    if (await pathExists(workbookPath)) throw new Error('Workbook already exists');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(resolvedTemplate);

    const bindings = await db.getMergeFieldBindingsByTemplate(TEMPLATE_BINDING_KEY);
    const mergeFields = await db.getMergeFields();
    const placeholderMap = new Map();
    (mergeFields || []).forEach(f => {
      const fieldKey = (f.field_key || '').toLowerCase();
      const placeholder = (f.placeholder || '').toLowerCase();
      const fieldSlug = normalizeTokenKey(fieldKey);
      const placeholderSlug = normalizeTokenKey(placeholder);
      if (fieldKey) placeholderMap.set(fieldKey, f.field_key);
      if (placeholder) placeholderMap.set(placeholder, f.field_key);
      if (fieldSlug) placeholderMap.set(fieldSlug, f.field_key);
      if (placeholderSlug) placeholderMap.set(placeholderSlug, f.field_key);
    });
    const placeholderKeys = collectPlaceholderKeys(workbook);
    const fieldKeySet = new Set((bindings || []).map(binding => binding.field_key).filter(Boolean));
    placeholderKeys.forEach(key => {
      const mapped = placeholderMap.get(String(key).toLowerCase()) || key;
      fieldKeySet.add(mapped);
    });
    const valueSources = await db.getMergeFieldValueSources(Array.from(fieldKeySet)) || {};
    await fillWorkbook(workbook, bindings, valueSources, context);
    replaceWorkbookPlaceholders(workbook, valueSources, context, placeholderMap);
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    sanitizeWorkbookValues(workbook);
    await workbook.xlsx.writeFile(workbookPath);

    // Reserve a document number (invoice/quote counter)
    const insert = await db.addDocument({
      business_id: businessId,
      doc_type: rawType,
      status: rawType === 'invoice' ? 'issued' : 'draft',
      total_amount: options.total_amount ?? null,
      balance_due: options.total_amount ?? null,
      due_date: options.due_date || null,
      client_name: context.client?.name || null,
      event_name: context.event?.event_name || null,
      event_date: context.event?.event_date || null,
      document_date: payload.document_date,
      definition_key: definitionKey
    });

    const number = insert?.number != null ? Number(insert.number) : null;
    // Build target PDF path; version if needed
    const baseName = path.basename(workbookPath, '.xlsx');
    let pdfBase = baseName;
    if (rawType === 'invoice' && number != null) {
      // If template encodes invoice code in filename, add INV-# suffix
      pdfBase = `${baseName} (INV-${number})`;
    } else if (rawType === 'quote' && number != null) {
      pdfBase = `${baseName} (Q-${number})`;
    }
    let pdfPath = path.join(directory, `${pdfBase}.pdf`);
    let n = 2;
    while (await pathExists(pdfPath)) {
      pdfPath = path.join(directory, `${pdfBase} (${n}).pdf`);
      n += 1;
      if (n > 1000) break;
    }

    // Determine stamp text for invoice; quotes typically do not carry a numeric stamp
    let stampCell = '';
    let stampText = '';
    if (rawType === 'invoice' && number != null) {
      try {
        const prefix = await readInvoicePrefixFromWorkbook(workbookPath);
        stampCell = 'E9';
        stampText = `${prefix}${number}`;
      } catch (_) {
        stampCell = 'E9';
        stampText = `INV-${number}`;
      }
    }

    // Export to PDF with optional stamping on the active sheet
    await saveWorkbookAsPdf(workbookPath, pdfPath, stampCell && stampText ? { stampCell, stampText } : {});

    // Update record with final path and status
    await db.updateDocumentStatus(insert.id, {
      file_path: pdfPath,
      status: 'issued',
      total_amount: options.total_amount ?? null,
      balance_due: options.total_amount ?? null,
      due_date: options.due_date || null
    });

    try {
      const cb = watcherCallbacks.get(businessId);
      if (typeof cb === 'function') cb({ businessId });
    } catch (_) {}

    return { ok: true, file_path: pdfPath, document_id: insert?.id || null, number };
  },
  // MCMS: Create a numbered invoice from a single Excel template with simple line items
  // Options: { business_id, definition_key='invoice_balance', client_override, line_items: [{ description, quantity, unit, rate, amount }], total_amount, due_date, document_date }
  createMCMSInvoice: async (options = {}) => {
    const businessId = Number(options.business_id ?? options.businessId);
    if (!Number.isInteger(businessId)) throw new Error('business_id is required');
    const rawItems = Array.isArray(options.line_items || options.items) ? (options.line_items || options.items) : [];
    const items = rawItems.map((it, idx) => {
      const desc = (it?.description || '').toString();
      const qty = Number(it?.quantity);
      const unit = (it?.unit || '').toString();
      const rate = Number(it?.rate);
      const amount = Number.isFinite(Number(it?.amount)) ? Number(it?.amount) : (Number.isFinite(qty) && Number.isFinite(rate) ? qty * rate : 0);
      return { description: desc, quantity: Number.isFinite(qty) ? qty : null, unit, rate: Number.isFinite(rate) ? rate : null, amount, sort_order: idx };
    }).filter(x => x.description || (x.amount != null && Number.isFinite(x.amount)));

    const definitionKey = options.definition_key || 'invoice_balance';
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');

    const def = await db.getDocumentDefinition(businessId, definitionKey);
    const templatePath = def?.template_path || options.template_path || options.templatePath;
    if (!templatePath) throw new Error('Template path is not configured for this document. Set it in Templates.');
    const resolvedTemplate = path.resolve(templatePath);
    await ensureFileAccessible(resolvedTemplate);

    // Compose context
    const payload = {
      business_id: businessId,
      client_override: options.client_override || {},
      total_amount: options.total_amount ?? (items.reduce((s, it) => s + (Number.isFinite(it.amount) ? it.amount : 0), 0) || null),
      balance_amount: options.total_amount ?? null,
      balance_due: options.total_amount ?? null,
      due_date: options.due_date || null,
      document_date: options.document_date || new Date().toISOString(),
      definition_key: definitionKey
    };
    const context = buildContext(payload, business);

    // Build naming and directories
    const naming = buildFileName(context, payload, { label: 'Invoice', key: definitionKey });
    const directory = buildOutputDirectory(business, context, payload, naming.folderName);
    await ensureDirectoryExists(directory);

    // Build workbook path
    const workbookPath = path.join(directory, `${naming.fileName}`);
    if (await pathExists(workbookPath)) throw new Error('Workbook already exists');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(resolvedTemplate);

    // Prepare placeholder mapping for invoice_code and client tokens
    const mergeFields = await db.getMergeFields();
    const placeholderMap = new Map();
    (mergeFields || []).forEach(f => {
      const fieldKey = (f.field_key || '').toLowerCase();
      const placeholder = (f.placeholder || '').toLowerCase();
      const fieldSlug = normalizeTokenKey(fieldKey);
      const placeholderSlug = normalizeTokenKey(placeholder);
      if (fieldKey) placeholderMap.set(fieldKey, f.field_key);
      if (placeholder) placeholderMap.set(placeholder, f.field_key);
      if (fieldSlug) placeholderMap.set(fieldSlug, f.field_key);
      if (placeholderSlug) placeholderMap.set(placeholderSlug, f.field_key);
    });

    // Also map invoice_code explicitly
    placeholderMap.set('invoice_code', 'invoice_code');

    const placeholderKeys = collectPlaceholderKeys(workbook);
    const fieldKeySet = new Set(['invoice_code', ...Array.from(placeholderKeys)]);
    const valueSources = await db.getMergeFieldValueSources(Array.from(fieldKeySet)) || {};
    // Inject a contextPath source for invoice_code -> context.invoiceCode
    valueSources['invoice_code'] = { source_type: 'contextPath', source_path: 'invoiceCode' };

    // Helper to fill the line items at an anchor {{items}}
    const writeLineItems = () => {
      let anchor = null;
      let anchorSheet = null;
      workbook.eachSheet(ws => {
        if (anchor) return;
        ws.eachRow((row, rowNumber) => {
          if (anchor) return;
          row.eachCell((cell, colNumber) => {
            if (anchor) return;
            const v = cell && typeof cell.value === 'string' ? cell.value.trim() : '';
            if (/^{{\s*items\s*}}$/i.test(v)) {
              anchor = { row: rowNumber, col: colNumber };
              anchorSheet = ws;
            }
          });
        });
      });
      if (!anchor || !anchorSheet) return;
      // Clear the anchor token
      try { anchorSheet.getCell(anchor.row, anchor.col).value = ''; } catch (_) {}

      // Copy simple style from anchor row if present
      const styleRow = anchorSheet.getRow(anchor.row);

      items.forEach((it, idx) => {
        const r = anchor.row + idx;
        const descCell = anchorSheet.getCell(r, anchor.col);
        const qtyCell = anchorSheet.getCell(r, anchor.col + 1);
        const unitCell = anchorSheet.getCell(r, anchor.col + 2);
        const rateCell = anchorSheet.getCell(r, anchor.col + 3);
        const amtCell = anchorSheet.getCell(r, anchor.col + 4);

        descCell.value = it.description || '';
        qtyCell.value = it.quantity != null && Number.isFinite(it.quantity) ? it.quantity : null;
        unitCell.value = it.unit || '';
        rateCell.value = it.rate != null && Number.isFinite(it.rate) ? it.rate : null;
        amtCell.value = it.amount != null && Number.isFinite(it.amount)
          ? it.amount
          : ((it.quantity != null && Number.isFinite(it.quantity) && it.rate != null && Number.isFinite(it.rate)) ? (it.quantity * it.rate) : null);

        // Apply number formats
        try { applyNumberFormat(rateCell, { data_type: 'number', format: 'currency' }); } catch (_) {}
        try { applyNumberFormat(amtCell, { data_type: 'number', format: 'currency' }); } catch (_) {}

        // Shallow style copy
        try {
          const srcDesc = styleRow.getCell(anchor.col);
          const srcQty = styleRow.getCell(anchor.col + 1);
          const srcUnit = styleRow.getCell(anchor.col + 2);
          const srcRate = styleRow.getCell(anchor.col + 3);
          const srcAmt = styleRow.getCell(anchor.col + 4);
          if (srcDesc && srcDesc.style) descCell.style = { ...srcDesc.style };
          if (srcQty && srcQty.style) qtyCell.style = { ...srcQty.style };
          if (srcUnit && srcUnit.style) unitCell.style = { ...srcUnit.style };
          if (srcRate && srcRate.style) rateCell.style = { ...srcRate.style };
          if (srcAmt && srcAmt.style) amtCell.style = { ...srcAmt.style };
        } catch (_) {}
      });
    };

    // Fill placeholders and items
    const issueDate = new Date();
    const dueDate = options.due_date || null;
    const amount = options.total_amount ?? (items.reduce((s, it) => s + (Number.isFinite(it.amount) ? it.amount : 0), 0) || null);

    // Reserve a document number now (invoice counter)
    const insert = await db.addDocument({
      business_id: businessId,
      doc_type: 'invoice',
      status: 'issued',
      total_amount: amount,
      balance_due: amount,
      due_date: dueDate || null,
      client_name: context.client?.name || null,
      event_name: null,
      event_date: null,
      document_date: payload.document_date,
      definition_key: definitionKey
    });
    const number = insert?.number != null ? Number(insert.number) : null;

    // Add invoice code to context for placeholder replacement
    const suffixCode = number != null ? `INV-${number}` : 'INV';
    const enrichedContext = { ...context, invoiceCode: suffixCode, issueDate, dueDate, totalAmount: amount };

    // Replace placeholders
    replaceWorkbookPlaceholders(workbook, valueSources, enrichedContext, placeholderMap);
    // Write items
    if (items.length) writeLineItems();
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    sanitizeWorkbookValues(workbook);
    await workbook.xlsx.writeFile(workbookPath);

    // Build PDF file name and export with stamping
    const baseName = path.basename(workbookPath, '.xlsx');
    let pdfBase = `${baseName} (INV-${number})`;
    let pdfPath = path.join(directory, `${pdfBase}.pdf`);
    let n = 2;
    while (await pathExists(pdfPath)) {
      pdfPath = path.join(directory, `${pdfBase} (${n}).pdf`);
      n += 1;
      if (n > 1000) break;
    }

    let stampCell = 'E9';
    let stampText = suffixCode;
    try {
      const prefix = await readInvoicePrefixFromWorkbook(workbookPath);
      stampText = `${prefix}${number}`;
    } catch (_) {}

    await saveWorkbookAsPdf(workbookPath, pdfPath, { stampCell, stampText });

    // Update record with final path
    await db.updateDocumentStatus(insert.id, {
      file_path: pdfPath,
      status: 'issued',
      total_amount: amount,
      balance_due: amount,
      due_date: dueDate || null
    });

    try {
      const cb = watcherCallbacks.get(businessId);
      if (typeof cb === 'function') cb({ businessId });
    } catch (_) {}

    return { ok: true, file_path: pdfPath, document_id: insert?.id || null, number };
  },
  exportWorkbookPdfs,
  deleteDocument,
  syncJobsheetOutputs,
  watchDocumentsFolder,
  unwatchDocumentsFolder,
  filterDocumentsByExistingFiles,
  listJobsheetDocuments,
  preflightPdfExport
  ,
  // Ensure a jobsheet's folder exists and return its path
  ensureJobsheetFolder: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
    const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;

    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) {
      throw new Error('Documents folder not configured for this business.');
    }

    let snapshot = (options.jobsheetSnapshot && typeof options.jobsheetSnapshot === 'object') ? { ...options.jobsheetSnapshot } : {};
    if ((!snapshot || Object.keys(snapshot).length === 0) && Number.isInteger(jobsheetId)) {
      try { snapshot = await db.getAhmenJobsheet(jobsheetId) || {}; } catch (_) { snapshot = {}; }
    }
    if (!snapshot.client_name && Number.isInteger(jobsheetId)) {
      snapshot.client_name = `Job ${jobsheetId}`;
    }
    if (!snapshot.event_date) {
      snapshot.event_date = new Date().toISOString().slice(0, 10);
    }

    const payload = { business_id: businessId, jobsheet_id: jobsheetId, jobsheet_snapshot: snapshot };
    const context = buildContext(payload, business);
    const folderPath = buildOutputDirectory(business, context, payload, 'Documents');
    await ensureDirectoryExists(folderPath);
    return { ok: true, folder_path: folderPath };
  },
  // Booking pack helpers
  getBookingPackPdfs,
  sendMailViaGraph,
  scheduleMailViaGraph,
  resolveTemplateDefaultAttachments,
  buildGigInfoHtml,
  buildPersonnelLogHtml,
  buildPersonnelLogText,
  // List files in a jobsheet's folder (non-recursive)
  listJobFolderFiles: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
    const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;
    if (!Number.isInteger(businessId)) throw new Error('businessId is required');
    const res = await module.exports.ensureJobsheetFolder({ businessId, jobsheetId });
    const folderPath = res?.folder_path || '';
    if (!folderPath) return [];
    let entries = [];
    try { entries = await fs.promises.readdir(folderPath, { withFileTypes: true }); } catch (_) { entries = []; }
    const files = [];
    for (const e of entries) {
      try {
        if (!e || !e.name || e.isDirectory()) continue;
        const p = path.join(folderPath, e.name);
        const st = await fs.promises.stat(p);
        files.push({ name: e.name, path: p, size: st.size, mtime: st.mtimeMs });
      } catch (_) {}
    }
    files.sort((a, b) => b.mtime - a.mtime);
    const filterPattern = options.extensionPattern ? new RegExp(options.extensionPattern, 'i') : null;
    if (filterPattern) {
      return files.filter(f => filterPattern.test(f.name || ''));
    }
    return files;
  },
  // Mail presets and signature (per business with global fallback)
  getMailPresets: async (options = {}) => {
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_presets_by_business || {};
      return map[businessId] || {};
    }
    return s.mail_presets || {};
  },
  saveMailPresets: async (options = {}) => {
    const presets = options.presets || {};
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_presets_by_business || {};
      map[businessId] = presets || {};
      s.mail_presets_by_business = map;
    } else {
      s.mail_presets = presets || {};
    }
    writeSettings(s);
    return { ok: true };
  },
  getMailSignature: async (options = {}) => {
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_signatures_by_business || {};
      const signature = map[businessId] || s.mail_signature || '';
      return { signature };
    }
    return { signature: s.mail_signature || '' };
  },
  saveMailSignature: async (options = {}) => {
    const signature = String(options.signature || '');
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_signatures_by_business || {};
      map[businessId] = signature;
      s.mail_signatures_by_business = map;
    } else {
      s.mail_signature = signature;
    }
    writeSettings(s);
    return { ok: true };
  },
  getDefaultMailTemplates: async (options = {}) => {
    const bName = String(options.businessName || '').trim();
    const labelFrom = (k, fallback) => fallback;
    const T = {
      enquiry_ack: {
        label: labelFrom('enquiry_ack', 'Enquiry acknowledgment'),
        subject: 'Thanks for your enquiry – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Thanks for getting in touch about your {{ event_type }} on {{ event_date }}.<br><br>I\'ll come back to you shortly with details and pricing.<br><br>Thanks,<br>'
      },
      quote: {
        label: labelFrom('quote', 'Quote'),
        subject: 'Quote – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Please find your quote attached for {{ event_type }} on {{ event_date }}.<br><br>If you have any questions or changes, just reply to this email.<br><br>Thanks,<br>'
      },
      booking_pack: {
        label: labelFrom('booking_pack', 'Booking pack'),
        subject: 'Booking pack – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Attached are your booking schedule and T&Cs, plus the deposit invoice.<br><br>Please review and let me know if anything needs updating.<br><br>Thanks,<br>'
      },
      invoice_deposit: {
        label: labelFrom('invoice_deposit', 'Deposit invoice'),
        subject: 'Deposit invoice – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Please find your deposit invoice attached.<br><br>Thanks,<br>'
      },
      invoice_balance: {
        label: labelFrom('invoice_balance', 'Balance invoice'),
        subject: 'Balance invoice – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Please find your balance invoice attached. The due date is {{ balance_due_date }}.<br><br>Thanks,<br>'
      },
      payment_reminder: {
        label: labelFrom('payment_reminder', 'Payment reminder'),
        subject: 'Payment reminder – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>A friendly reminder about the outstanding balance for {{ event_type }} on {{ event_date }}. The due date is {{ balance_due_date }}.<br><br>If you\'ve already sent payment, please ignore this.<br><br>Thanks,<br>'
      },
      thank_you: {
        label: labelFrom('thank_you', 'Thank you'),
        subject: 'Thank you – {{ client_name }} – {{ event_date }}',
        body: 'Hi {{ client_first_name|there }},<br><br>Thank you again for having us at your {{ event_type }} on {{ event_date }} — it was a pleasure.<br><br>All the best,<br>'
      }
    };
    return T;
  },
  // Gig Info presets (dress code and repertoire)
  getGigInfoPresets: async (options = {}) => {
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    const norm = (v) => Array.isArray(v) ? v.filter(x => typeof x === 'string' && x.trim()).map(String) : [];
    if (Number.isInteger(businessId)) {
      const map = s.gig_info_presets_by_business || {};
      const entry = map[businessId] || {};
      return {
        dress_codes: norm(entry.dress_codes),
        repertoire: norm(entry.repertoire)
      };
    }
    const def = s.gig_info_presets || {};
    return {
      dress_codes: norm(def.dress_codes),
      repertoire: norm(def.repertoire)
    };
  },
  saveGigInfoPreset: async (options = {}) => {
    const kindRaw = String(options.kind || '').toLowerCase();
    const kind = kindRaw === 'dress_code' || kindRaw === 'dress' ? 'dress_codes' : (kindRaw === 'repertoire' || kindRaw === 'setlist' ? 'repertoire' : null);
    const value = String(options.value || '').trim();
    const businessId = Number(options.businessId ?? options.business_id);
    if (!kind) throw new Error('Invalid preset kind');
    if (!value) throw new Error('Preset value is empty');
    const s = readSettings();
    const cap = Number.isInteger(options.limit) ? options.limit : 20;
    const pushUniqueFront = (arr, v) => {
      const list = Array.isArray(arr) ? arr.slice() : [];
      const existingIndex = list.findIndex(x => String(x) === v);
      if (existingIndex >= 0) list.splice(existingIndex, 1);
      list.unshift(v);
      if (cap > 0 && list.length > cap) list.length = cap;
      return list;
    };
    if (Number.isInteger(businessId)) {
      const map = s.gig_info_presets_by_business || {};
      const entry = map[businessId] || {};
      entry[kind] = pushUniqueFront(entry[kind], value);
      map[businessId] = entry;
      s.gig_info_presets_by_business = map;
    } else {
      const def = s.gig_info_presets || {};
      def[kind] = pushUniqueFront(def[kind], value);
      s.gig_info_presets = def;
    }
    writeSettings(s);
    return { ok: true };
  },
  deleteGigInfoPreset: async (options = {}) => {
    const kindRaw = String(options.kind || '').toLowerCase();
    const kind = kindRaw === 'dress_code' || kindRaw === 'dress' ? 'dress_codes' : (kindRaw === 'repertoire' || kindRaw === 'setlist' ? 'repertoire' : null);
    const value = String(options.value || '').trim();
    const businessId = Number(options.businessId ?? options.business_id);
    if (!kind) throw new Error('Invalid preset kind');
    if (!value) throw new Error('Preset value is empty');
    const s = readSettings();
    const remove = (arr, v) => (Array.isArray(arr) ? arr.filter(x => String(x) !== v) : []);
    if (Number.isInteger(businessId)) {
      const map = s.gig_info_presets_by_business || {};
      const entry = map[businessId] || {};
      entry[kind] = remove(entry[kind], value);
      map[businessId] = entry;
      s.gig_info_presets_by_business = map;
    } else {
      const def = s.gig_info_presets || {};
      def[kind] = remove(def[kind], value);
      s.gig_info_presets = def;
    }
    writeSettings(s);
    return { ok: true };
  },
  renameGigInfoPreset: async (options = {}) => {
    const kindRaw = String(options.kind || '').toLowerCase();
    const kind = kindRaw === 'dress_code' || kindRaw === 'dress' ? 'dress_codes' : (kindRaw === 'repertoire' || kindRaw === 'setlist' ? 'repertoire' : null);
    const from = String(options.from || '').trim();
    const to = String(options.to || '').trim();
    const businessId = Number(options.businessId ?? options.business_id);
    if (!kind) throw new Error('Invalid preset kind');
    if (!from || !to) throw new Error('Both from and to are required');
    const s = readSettings();
    const replace = (arr, a, b) => {
      const list = Array.isArray(arr) ? arr.slice() : [];
      const idx = list.findIndex(x => String(x) === a);
      if (idx === -1) return list;
      // Remove any existing identical target, then set new at same index
      const filtered = list.filter(x => String(x) !== b);
      filtered.splice(idx, 1, b);
      return filtered;
    };
    if (Number.isInteger(businessId)) {
      const map = s.gig_info_presets_by_business || {};
      const entry = map[businessId] || {};
      entry[kind] = replace(entry[kind], from, to);
      map[businessId] = entry;
      s.gig_info_presets_by_business = map;
    } else {
      const def = s.gig_info_presets || {};
      def[kind] = replace(def[kind], from, to);
      s.gig_info_presets = def;
    }
    writeSettings(s);
    return { ok: true };
  },
  // Mail templates (subject/body only), per business
  getMailTemplates: async (options = {}) => {
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    const map = s.mail_templates_by_business || {};
    if (Number.isInteger(businessId)) return map[businessId] || {};
    return s.mail_templates || {};
  },
  getMailTemplateTombstones: async (options = {}) => {
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    const map = s.mail_template_tombstones_by_business || {};
    if (Number.isInteger(businessId)) return Array.isArray(map[businessId]) ? map[businessId] : [];
    const arr = s.mail_template_tombstones;
    return Array.isArray(arr) ? arr : [];
  },
  saveMailTemplates: async (options = {}) => {
    const templates = options.templates || {};
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_templates_by_business || {};
      map[businessId] = templates || {};
      s.mail_templates_by_business = map;
    } else {
      s.mail_templates = templates || {};
    }
    writeSettings(s);
    return { ok: true };
  },
  deleteMailTemplate: async (options = {}) => {
    const keyRaw = options.key != null ? String(options.key) : '';
    const key = keyRaw.toLowerCase().trim();
    if (!key) return { ok: false, message: 'Key required' };
    const s = readSettings();
    const businessId = Number(options.businessId ?? options.business_id);
    if (Number.isInteger(businessId)) {
      const map = s.mail_templates_by_business || {};
      const forBiz = map[businessId] || {};
      if (forBiz[key]) {
        delete forBiz[key];
        map[businessId] = forBiz;
        s.mail_templates_by_business = map;
      }
      const tombMap = s.mail_template_tombstones_by_business || {};
      const list = Array.isArray(tombMap[businessId]) ? tombMap[businessId] : [];
      if (!list.includes(key)) list.push(key);
      tombMap[businessId] = list;
      s.mail_template_tombstones_by_business = tombMap;
    } else {
      const map = s.mail_templates || {};
      if (map[key]) delete map[key];
      s.mail_templates = map;
      const tombList = Array.isArray(s.mail_template_tombstones) ? s.mail_template_tombstones : [];
      if (!tombList.includes(key)) tombList.push(key);
      s.mail_template_tombstones = tombList;
    }
    writeSettings(s);
    return { ok: true };
  },
  // (legacy "other files" helpers removed)
  extractJobsheetDataFromFolder,
  indexInvoicesFromFilenames: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) {
      throw new Error('Documents folder not configured for this business.');
    }
    const root = path.resolve(business.save_path);

    // Collect candidate PDFs with (INV-###) in their filename
    async function walk(dir, depth = 0, maxDepth = 6) {
      let entries = [];
      try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch (_) { return []; }
      const results = [];
      for (const e of entries) {
        if (e.name.startsWith('.')) continue;
        const full = path.join(dir, e.name);
        if (e.isDirectory()) {
          if (depth < maxDepth) {
            // eslint-disable-next-line no-await-in-loop
            const sub = await walk(full, depth + 1, maxDepth);
            results.push(...sub);
          }
          continue;
        }
        const lower = e.name.toLowerCase();
        if (!lower.endsWith('.pdf')) continue;
        if (!/inv[\-\s]?\d+/i.test(e.name)) continue;
        results.push(full);
      }
      return results;
    }

    const files = await walk(root);
    if (!files.length) return { imported: 0 };

    // Preload jobsheets for matching
    let sheets = [];
    try { sheets = await db.getAhmenJobsheets({ businessId }); } catch (_) { sheets = []; }
    const norm = s => (s || '').toString().trim().toLowerCase();

    const matchJobsheet = (filePath) => {
      try {
        const dirBase = path.basename(path.dirname(filePath));
        const parts = dirBase.split(' - ');
        const dateStr = parts[0] || '';
        const client = parts[1] || '';
        const dateOk = /^\d{4}-\d{2}-\d{2}$/.test(dateStr) ? dateStr : '';
        const cand = sheets.find(js => (!dateOk || norm(js.event_date) === norm(dateOk)) && (client ? norm(js.client_name) === norm(client) : true));
        return cand || null;
      } catch (_) { return null; }
    };

    let imported = 0;
    for (const pdfPath of files) {
      try {
        // Skip if invoice already recorded
        // eslint-disable-next-line no-await-in-loop
        const existing = await db.getDocumentByFilePath(businessId, pdfPath);
        if (existing && String(existing.doc_type || '').toLowerCase() === 'invoice') continue;

        const base = path.basename(pdfPath);
        const numMatch = base.match(/inv[\-\s]?(\d+)/i);
        const number = numMatch ? Number(numMatch[1]) : null;
        if (number == null || !Number.isInteger(number)) continue;

        const nameLower = base.toLowerCase();
        const variant = nameLower.includes('deposit') ? 'deposit' : (nameLower.includes('balance') ? 'balance' : null);
        const js = matchJobsheet(pdfPath);

        const payload = {
          business_id: businessId,
          jobsheet_id: js?.jobsheet_id || null,
          doc_type: 'invoice',
          number,
          status: 'issued',
          total_amount: variant === 'deposit' ? (js?.deposit_amount ?? null) : (js?.balance_amount ?? null),
          balance_due: variant === 'balance' ? (js?.balance_amount ?? null) : (js?.deposit_amount ?? js?.balance_amount ?? null),
          due_date: variant === 'balance' ? (js?.balance_due_date ?? null) : (js?.event_date ?? null),
          file_path: pdfPath,
          client_name: js?.client_name || null,
          event_name: js?.event_type || null,
          event_date: js?.event_date || null,
          document_date: js?.updated_at || new Date().toISOString(),
          definition_key: null,
          invoice_variant: variant
        };
        try {
          // eslint-disable-next-line no-await-in-loop
          await db.addDocument(payload);
          imported += 1;
        } catch (insErr) {
          // Number conflict: ensure last_invoice_number sync but skip create
          // eslint-disable-next-line no-console
          console.warn('Invoice import skipped', base, insErr?.message || insErr);
        }
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('Failed to import invoice pdf', pdfPath, err);
      }
    }

    try {
      const maxNum = await db.getMaxInvoiceNumber(businessId);
      const last = Number.isInteger(Number(maxNum)) ? Number(maxNum) : 0;
      await db.setLastInvoiceNumber(businessId, last);
    } catch (_) {}

    return { imported };
  },
  computeFinderInvoiceMax: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) {
      throw new Error('Documents folder not configured for this business.');
    }
    const root = path.resolve(business.save_path);

    async function walk(dir, depth = 0, maxDepth = 6) {
      let entries = [];
      try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch (_) { return []; }
      const results = [];
      for (const e of entries) {
        if (e.name.startsWith('.')) continue;
        const full = path.join(dir, e.name);
        if (e.isDirectory()) {
          if (depth < maxDepth) {
            // eslint-disable-next-line no-await-in-loop
            const sub = await walk(full, depth + 1, maxDepth);
            results.push(...sub);
          }
          continue;
        }
        const lower = e.name.toLowerCase();
        if (!lower.endsWith('.pdf')) continue;
        const m = e.name.match(/inv[\-\s]?(\d+)/i);
        if (m && m[1]) {
          const num = Number(m[1]);
          if (Number.isInteger(num)) results.push(num);
        }
      }
      return results;
    }

    const numbers = await walk(root);
    const max = numbers.length ? Math.max(...numbers) : 0;
    return { max };
  },
  cleanOrphanDocuments: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    if (!Number.isInteger(businessId)) {
      throw new Error('businessId is required');
    }
    const performDelete = options.delete === true || options.remove === true;
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) {
      throw new Error('Documents folder not configured for this business.');
    }
    const rootPath = path.resolve(business.save_path);

    const docs = await db.getDocuments({ businessId });
    const existing = await filterDocumentsByExistingFiles(docs, { includeMissing: true });

    // Clear DB for missing files
    let clearedMissing = 0;
    for (const doc of existing) {
      const hasFile = doc?.file_available === true;
      if (!hasFile && doc?.file_path) {
        try { await db.clearDocumentPath(businessId, doc.file_path); clearedMissing += 1; } catch (_err) {}
      }
    }

    // Group by folder and find orphan PDFs (no matching workbook base)
    const byDir = new Map();
    existing.forEach(doc => {
      const fp = doc?.file_path || '';
      if (!fp) return;
      const dir = path.dirname(fp);
      const arr = byDir.get(dir) || [];
      arr.push(doc);
      byDir.set(dir, arr);
    });

    const orphanRecords = [];
    for (const [dir, list] of byDir.entries()) {
      const workbookBases = new Set(
        list
          .filter(d => (d?.file_path || '').toLowerCase().endsWith('.xlsx'))
          .map(d => path.basename(d.file_path, path.extname(d.file_path)))
      );
      for (const d of list) {
        const fp = d?.file_path || '';
        if (!fp.toLowerCase().endsWith('.pdf')) continue;
        const base = path.basename(fp, path.extname(fp));
        if (!workbookBases.has(base)) {
          orphanRecords.push(d);
        }
      }
    }

    let deleted = 0;
    const records = [];
    for (const d of orphanRecords) {
      if (!d || !d.file_path) continue;
      if (d.is_locked) { records.push({ file_path: d.file_path, action: 'locked' }); continue; }
      if (!isSubPath(rootPath, d.file_path)) { records.push({ file_path: d.file_path, action: 'outside-root' }); continue; }
      if (performDelete) {
        try { await db.deleteDocumentByPath({ businessId, absolutePath: d.file_path }); deleted += 1; records.push({ file_path: d.file_path, action: 'deleted' }); } catch (_err) { records.push({ file_path: d.file_path, action: 'failed' }); }
      } else {
        records.push({ file_path: d.file_path, action: 'orphan' });
      }
    }

    return { ok: true, cleared_missing: clearedMissing, orphan_count: orphanRecords.length, deleted, records };
  }
};

const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
let chokidar = null;
try { chokidar = require('chokidar'); } catch (_err) { chokidar = null; }
const ExcelJS = require('exceljs');
const { PDFDocument, StandardFonts, TextAlignment, rgb } = require('pdf-lib');
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

const SCHEDULED_WORKER_OVERRIDE = String(process.env.AHMEN_SCHEDULED_WORKER || '').trim() === '1';
const SCHEDULED_WORKER_ENABLED = (() => {
  if (!IS_MAIN_PROCESS) return false;
  if (SCHEDULED_WORKER_OVERRIDE) return true;
  const argv = Array.isArray(process.argv) ? process.argv.join(' ') : '';
  if (argv.includes('--background') || argv.includes('--helper')) return true;
  const execPath = process.execPath || '';
  if (execPath.includes(`${path.sep}LoginItems${path.sep}`)) return true;
  if (/ahmen reminders/i.test(execPath)) return true;
  return false;
})();

let scheduledMailWorkerStarted = false;
let scheduledMailWorkerExecuting = false;
let ElectronBrowserWindow = null;

async function syncFinderInvoiceCounter(businessId) {
  const id = Number(businessId);
  if (!Number.isInteger(id)) return null;
  try {
    if (typeof module.exports.computeFinderInvoiceMax !== 'function') return null;
    const business = await db.getBusinessById(id);
    const currentRaw = business?.last_invoice_number;
    const current = Number.isInteger(Number(currentRaw)) ? Number(currentRaw) : 0;
    const result = await module.exports.computeFinderInvoiceMax({ businessId: id });
    const rawMax = result && result.max != null ? Number(result.max) : 0;
    const max = Number.isInteger(rawMax) && rawMax >= 0 ? rawMax : 0;
    const next = Math.max(current, max);
    if (next !== current) {
      await db.setLastInvoiceNumber(id, next);
    }
    return next;
  } catch (_err) {
    return null;
  }
}

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

function splitAddressLines(raw) {
  if (!raw) return [];
  return String(raw)
    .split(/\r?\n+/)
    .map(line => line.trim())
    .filter(Boolean);
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
  if (!SCHEDULED_WORKER_ENABLED) return;
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
    try {
      await sendInternalNotice({
        businessId: entry.business_id,
        jobsheetId: entry.jobsheet_id,
        subject: `Scheduled email sent · ${entry.subject || 'Email'}`,
        body: `Scheduled email sent to ${escapeHtml(entry.to_address || '')}.<br>Subject: ${escapeHtml(entry.subject || '')}`
      });
    } catch (_) {}
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

const MAIL_TOKEN_REGEX = /{{\s*([a-zA-Z0-9_.-]+)(?:\|([^}]+))?\s*}}/g;
const SIG_START = '<!--__IM_SIG_START__-->';
const SIG_END = '<!--__IM_SIG_END__-->';

function parseIsoDateParts(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return { year, month, day };
}

function formatDateKey(parts) {
  if (!parts) return '';
  const pad = (val) => String(val).padStart(2, '0');
  return `${parts.year}-${pad(parts.month)}-${pad(parts.day)}`;
}

function normalizeDateKey(value) {
  const parts = parseIsoDateParts(value);
  return formatDateKey(parts);
}

function addDaysISO(dateStr, offset) {
  if (!dateStr) return '';
  const base = new Date(dateStr);
  if (Number.isNaN(base.valueOf())) return '';
  base.setDate(base.getDate() + offset);
  return base.toISOString().slice(0, 10);
}

function addMonthsISO(dateStr, offset) {
  if (!dateStr) return '';
  const parts = parseIsoDateParts(dateStr);
  if (!parts) return '';
  const date = new Date(parts.year, parts.month - 1 + Number(offset || 0), parts.day);
  if (Number.isNaN(date.valueOf())) return '';
  return date.toISOString().slice(0, 10);
}

function buildMailTokenMap(snapshot = {}) {
  const js = snapshot || {};
  const firstName = (() => {
    const raw = String(js.client_name || '').trim();
    if (!raw) return '';
    return raw.split(/\s+/)[0] || '';
  })();
  const money = (val) => {
    const num = Number(val);
    if (!Number.isFinite(num)) return '';
    return formatCurrencyGBP(num);
  };

  return {
    client_name: js.client_name || '',
    client_first_name: firstName,
    client_email: js.client_email || '',
    event_type: js.event_type || '',
    event_date: formatDisplayDate(js.event_date || ''),
    balance_due_date: formatDisplayDate(js.balance_due_date || ''),
    balance_reminder_date: formatDisplayDate(js.balance_reminder_date || ''),
    balance_amount: money(js.balance_amount),
    total_amount: money(js.total_amount),
    today: formatDisplayDate(new Date())
  };
}

function renderMailTemplate(template, tokenMap = {}) {
  if (!template) return '';
  return String(template).replace(MAIL_TOKEN_REGEX, (_match, key, fallback) => {
    const normalizedKey = String(key || '').trim().toLowerCase();
    const value = tokenMap[normalizedKey];
    if (value != null && value !== '') return String(value);
    return fallback != null ? String(fallback) : '';
  });
}

function appendSignatureHtml(bodyHtml, signatureHtml) {
  const trimmedBody = (bodyHtml || '').trim();
  if (!signatureHtml) return trimmedBody;
  const wrappedSig = `${SIG_START}${signatureHtml}${SIG_END}`;
  if (!trimmedBody) return wrappedSig;
  if (/(<br\s*\/?>|<\/p>)$/i.test(trimmedBody)) {
    return `${trimmedBody}${wrappedSig}`;
  }
  return `${trimmedBody}<br><br>${wrappedSig}`;
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

  const populateAddress = (target, combinedKey, prefix) => {
    if (!target || !combinedKey || !prefix) return;
    const combined = target[combinedKey];
    if (!combined) return;
    const parts = splitAddressLines(combined);
    const assign = (key, val) => {
      if (target[key] == null || target[key] === '') {
        target[key] = val;
      }
    };
    assign(`${prefix}_address1`, parts[0] || '');
    assign(`${prefix}_address2`, parts[1] || '');
    assign(`${prefix}_address3`, parts[2] || '');
    assign(`${prefix}_town`, parts[3] || '');
    assign(`${prefix}_postcode`, parts[4] || '');
  };

  populateAddress(jobsheet, 'client_address', 'client');
  populateAddress(jobsheet, 'venue_address', 'venue');

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
              } else if (typeof resolved === 'string') {
                // Allow non-date strings like "On receipt"
                cell.value = resolved;
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
                  } else if (typeof resolved === 'string') {
                    cell.value = resolved; // keep plain text like "On receipt"
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
              } else if (typeof resolved === 'string') {
                cell.value = resolved;
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
    // Ensure we always have a reference to the active sheet for setup
    '-e', 'set theSheet to active sheet of wb',
    // If stamping info was provided, update the cell value on the target or active sheet before export
    '-e', 'if stampCell is not missing value and stampText is not missing value then',
    '-e', 'try',
    // theSheet already set; may be reassigned below if a specific sheet is found
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
    // Enforce page setup: one page wide, unlimited tall, and constrain print area to used range
    '-e', 'try',
    '-e', 'set ps to page setup object of theSheet',
    '-e', 'set zoom of ps to false',
    '-e', 'set fit to pages wide of ps to 1',
    '-e', 'set fit to pages tall of ps to false',
    '-e', 'try',
    '-e', 'set rng to used range of theSheet',
    '-e', 'set print area of ps to (get address of rng)',
    '-e', 'end try',
    '-e', 'end try',
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

function normalizeProductionItemsServer(raw) {
  let items = [];
  if (Array.isArray(raw)) {
    items = raw;
  } else if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) items = parsed;
    } catch (_) {
      items = [];
    }
  }
  return items
    .map((item, index) => {
      if (!item) return null;
      return {
        id: item.id != null ? String(item.id) : `prod-${index}`,
        name: item.name != null ? String(item.name) : '',
        description: item.description != null ? String(item.description) : '',
        cost: item.cost != null ? String(item.cost) : '',
        markup: item.markup != null ? String(item.markup) : '',
        notes: item.notes != null ? String(item.notes) : ''
      };
    })
    .filter(it => it && (it.name || it.description || it.cost || it.notes));
}

function calculateProductionItemTotalServer(item) {
  if (!item) return 0;
  const base = Number(String(item.cost || '').replace(/[^0-9.\-]+/g, ''));
  const markupPct = Number(String(item.markup || '').replace(/[^0-9.\-]+/g, ''));
  const safeBase = Number.isFinite(base) ? base : 0;
  const fraction = Number.isFinite(markupPct) ? markupPct / 100 : 0;
  const total = safeBase + safeBase * fraction;
  return Number.isFinite(total) ? total : 0;
}

function formatProductionLines(items = []) {
  const labels = normalizeProductionItemsServer(items).map((item) => {
    const label = (item.description || '').trim() || (item.name || '').trim();
    const amount = calculateProductionItemTotalServer(item);
    const amountLabel = Number.isFinite(amount) && amount > 0 ? formatCurrencyGBP(amount) : '';
    const amountPart = amountLabel ? (label ? `(${amountLabel})` : amountLabel) : '';
    const notes = item.notes ? `(${item.notes})` : '';
    // Only show client-facing description (no supplier/company)
    return [label, amountPart, notes].filter(Boolean).join(' ').trim();
  }).filter(Boolean);

  if (!labels.length) return ['', ''];

  const maxLen = 90;
  let first = '';
  let second = '';
  labels.forEach((entry) => {
    const candidate = first ? `${first}; ${entry}` : entry;
    if (!first || candidate.length <= maxLen) {
      first = candidate;
    } else {
      second = second ? `${second}; ${entry}` : entry;
    }
  });

  return [first, second];
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

async function createPdfDocument(payload = {}) {
  const businessId = Number(payload.business_id ?? payload.businessId);
  if (!Number.isInteger(businessId)) {
    throw new Error('business_id is required to generate documents.');
  }

  const templatePath = resolvePath(payload.template_path || payload.templatePath);
  await ensureFileAccessible(templatePath);

  const business = await db.getBusinessById(businessId);
  if (!business) {
    throw new Error('Business record not found.');
  }

  const context = buildContext(payload, business);
  const definitionKey = payload.definition_key || payload.document_definition_key || 'workbook';
  const definition = await db.getDocumentDefinition(businessId, definitionKey);
  const naming = buildFileName(context, payload, definition);
  const directory = buildOutputDirectory(business, context, payload, naming.folderName);
  await fs.promises.mkdir(directory, { recursive: true });

  const docTypeRaw = (payload.doc_type || payload.type || definition?.doc_type || 'workbook').toLowerCase() || 'workbook';

  const js = context.jobsheet || {};
  const client = context.client || {};
  const event = context.event || {};
  const derived = context.context || {};
  const pricing = context.pricing || {};

  const parseAmount = (val) => {
    if (val === null || val === undefined || val === '') return null;
    const num = Number(String(val).replace(/[^0-9.\-]+/g, ''));
    return Number.isFinite(num) ? Math.round(num * 100) / 100 : null;
  };
  const currencyOrEmpty = (val) => {
    const num = parseAmount(val);
    return num === null ? '' : formatCurrencyGBP(num);
  };

  const totalAmount = parseAmount(derived.totalAmount ?? js.pricing_total ?? js.total_amount ?? payload.total_amount);
  const depositAmount = parseAmount(derived.depositAmount ?? js.deposit_amount ?? payload.deposit_amount);
  const balanceAmount = parseAmount(derived.balanceAmount ?? js.balance_amount ?? payload.balance_amount ?? derived.balanceDue);
  const ahmenFee = parseAmount(js.ahmen_fee);
  const productionFees = parseAmount(derived.productionFees ?? js.production_fees ?? js.pricing_production_total);
  const vatEnabled = Boolean(js.vat_enabled);
  const vatAmount = parseAmount(js.vat_amount);
  const vatRate = 0.2;
  const quoteSubtotal = Math.max((ahmenFee || 0) + (productionFees || 0), 0);
  const computedVat = Math.max(quoteSubtotal * vatRate, 0);
  const effectiveVat = vatEnabled ? (vatAmount != null && vatAmount !== '' ? vatAmount : computedVat) : 0;
  const quoteTotal = Math.max(quoteSubtotal + (vatEnabled ? effectiveVat : 0), 0);
  const extraFees = 0; // legacy field removed from UI; keep at zero for PDFs

  const invoiceVariant = payload.invoice_variant || definition?.invoice_variant || null;
  const invoiceAmount = (() => {
    if (docTypeRaw !== 'invoice') return null;
    if (invoiceVariant === 'deposit' && Number.isFinite(depositAmount)) return depositAmount;
    if (invoiceVariant === 'balance' && Number.isFinite(balanceAmount)) return balanceAmount;
    return Number.isFinite(totalAmount) ? totalAmount : null;
  })();
  const invoiceTotal = Number.isFinite(invoiceAmount) ? invoiceAmount : quoteTotal;
  let reservedInvoiceId = null;
  let reservedInvoiceNumber = null;
  if (docTypeRaw === 'invoice') {
    await syncFinderInvoiceCounter(businessId);
    const documentDate = payload.document_date || new Date().toISOString();
    try {
      const inserted = await db.addDocument({
        business_id: businessId,
        jobsheet_id: payload.jobsheet_id || null,
        doc_type: 'invoice',
        status: 'issued',
        total_amount: Number.isFinite(invoiceTotal) ? invoiceTotal : 0,
        balance_due: Number.isFinite(invoiceTotal) ? invoiceTotal : 0,
        due_date: payload.balance_due_date ?? payload.due_date ?? null,
        file_path: null,
        client_name: js.client_name || client.name || null,
        event_name: js.event_name || js.event_type || null,
        event_date: js.event_date || event.event_date || null,
        document_date: documentDate,
        definition_key: definitionKey,
        invoice_variant: invoiceVariant
      });
      reservedInvoiceId = inserted?.id || null;
      reservedInvoiceNumber = inserted?.number != null ? Number(inserted.number) : null;
    } catch (err) {
      console.warn('Failed to reserve invoice number', err);
    }
  }

  const ext = path.extname(templatePath) || '.pdf';
  const baseStemRaw = naming.fileName ? naming.fileName.replace(/\.xlsx$/i, '').replace(/\.[^.]+$/, '') : '';
  const baseStem = baseStemRaw || 'Document';
  const invoiceSuffix = reservedInvoiceNumber != null ? ` (INV-${reservedInvoiceNumber})` : '';
  const stemWithSuffix = docTypeRaw === 'invoice' ? `${baseStem}${invoiceSuffix}` : baseStem;
  let targetPath = path.join(directory, `${stemWithSuffix}${ext}`);
  let counter = 2;
  while (await pathExists(targetPath)) {
    try {
      const existing = await db.getDocumentByFilePath(businessId, targetPath);
      if (existing && existing.is_locked) {
        throw new Error('Document is locked');
      }
    } catch (_err) {}
    targetPath = path.join(directory, `${stemWithSuffix} (${counter})${ext}`);
    counter += 1;
    if (counter > 200) break;
  }

  const templateBytes = await fs.promises.readFile(templatePath);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const form = pdfDoc.getForm();

  // Font selection: default Helvetica 10pt for non-quotes; quotes keep template styling
  const baseFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const applyFieldStyle = (field) => {
    if (!field) return;
    if (docTypeRaw === 'quote') return; // keep template-defined appearance for quotes
    try {
      field.setFont(baseFont);
      field.setFontSize(10);
    } catch (_) {}
  };

  const setText = (name, value, align) => {
    try {
      const field = form.getTextField(name);
      if (field) {
        field.setText(value == null ? '' : String(value));
        if (align) {
          try { field.setAlignment(align); } catch (_) {}
        }
        applyFieldStyle(field);
        // Some quote fields (notably VAT) need fresh appearances to render values
        const upper = name.toUpperCase();
        const needsExplicitAppearance = upper.startsWith('VAT') || upper === 'DATE_TODAY' || upper === 'TODAY_DATE';
        if (needsExplicitAppearance) {
          try {
            // Preserve template font if available; fallback to baseFont
            const da = field.acroField.getDefaultAppearance();
            if (da) {
              field.defaultUpdateAppearances(da);
            } else {
              field.setFont(baseFont);
              field.setFontSize(9);
              field.setTextColor(rgb(0, 0, 0));
              field.setAlignment(align || TextAlignment.Right);
              field.updateAppearances(baseFont);
            }
          } catch (_) {}
        }
      }
    } catch (_err) {}
  };
  const setCheck = (name, value) => {
    try {
      const field = form.getCheckBox(name);
      if (!field) return;
      if (value) field.check();
      else field.uncheck();
      applyFieldStyle(field);
    } catch (_err) {}
  };

  const [prodLine1, prodLine2] = formatProductionLines(js.pricing_production_items || pricing.pricing_production_items || []);
  const startTime = formatTimeLabel(js.event_start || event.event_start || '', null);
  const endTime = formatTimeLabel(null, js.event_end || event.event_end || '');
  const eventDateDisplay = formatDisplayDate(js.event_date || event.event_date || '');
  const balanceDateDisplay = formatDisplayDate(js.balance_due_date || derived.balanceDate || payload.balance_due_date || payload.due_date || '');
  const clientLines = (() => {
    const combined = js.client_address || '';
    const parts = splitAddressLines(combined);
    if (parts.length) return parts;
    return [
      js.client_address1,
      js.client_address2,
      js.client_address3,
      [js.client_town, js.client_postcode].filter(Boolean).join(' ')
    ].filter(Boolean);
  })();
  const specialRaw = String(js.special_conditions || '').trim();
  const specialLines = specialRaw
    ? specialRaw.split(/\r?\n/).map(s => s.trim()).filter(Boolean)
    : [];
  while (specialLines.length > 3) {
    const overflow = specialLines.splice(3).join(' ');
    specialLines[2] = `${specialLines[2]} ${overflow}`.trim();
  }
  const usesInvoiceVariant = docTypeRaw === 'invoice' && (invoiceVariant === 'deposit' || invoiceVariant === 'balance');
  const fullVat = vatEnabled ? effectiveVat : 0;
  let displaySubtotal = quoteSubtotal;
  let displayVat = fullVat;
  let displayTotal = quoteTotal;
  let displayDepositAmount = depositAmount;
  let displayBalanceAmount = balanceAmount;
  if (docTypeRaw === 'invoice') {
    displayTotal = Number.isFinite(invoiceTotal) ? invoiceTotal : quoteTotal;
    if (usesInvoiceVariant) {
      const ratio = quoteTotal > 0 ? displayTotal / quoteTotal : 0;
      displayVat = vatEnabled ? Math.round(fullVat * ratio * 100) / 100 : 0;
      displaySubtotal = vatEnabled ? Math.max(displayTotal - displayVat, 0) : displayTotal;
    } else {
      displayVat = vatEnabled ? fullVat : 0;
      displaySubtotal = vatEnabled ? Math.max(displayTotal - displayVat, 0) : displayTotal;
    }
    if (invoiceVariant === 'deposit') {
      displayDepositAmount = displaySubtotal;
    } else if (invoiceVariant === 'balance') {
      displayBalanceAmount = displaySubtotal;
    }
  } else if (docTypeRaw === 'quote' && vatEnabled) {
    const netRatio = quoteTotal > 0 ? quoteSubtotal / quoteTotal : 0;
    if (Number.isFinite(depositAmount) && netRatio > 0) {
      displayDepositAmount = Math.round(depositAmount * netRatio * 100) / 100;
    }
    if (Number.isFinite(balanceAmount) && netRatio > 0) {
      displayBalanceAmount = Math.round(balanceAmount * netRatio * 100) / 100;
    }
  }

  setText('CLIENT_NAME', js.client_name || client.name || '');
  setText('CLIENT_EMAIL', js.client_email || client.email || '');
  setText('CLIENT_PHONE', js.client_phone || client.phone || '');
  setText('CLIENT_ADDRESS1', clientLines[0] || '');
  setText('CLIENT_ADDRESS2', clientLines[1] || '');
  setText('CLIENT_ADDRESS3', clientLines[2] || '');
  setText('CLIENT_ADDRESS4', clientLines[3] || '');
  setText('CLIENT_NAME_es_:signer', js.client_name || client.name || '');
  setText('CLIENT_SIGN_es_:signer:signature', '');
  setText('EVENT_DATE', eventDateDisplay);
  setText('EVENT_START', startTime);
  setText('EVENT_END', endTime);
  setText('EVENT_TYPE', js.event_type || event.event_type || '');
  setText('SERVICE_TYPE', js.service_types || '');
  setText('SPECIALIST_SINGERS', js.specialist_singers || '');
  setText('VENUE_NAME', js.venue_name || '');
  const venueLines = (() => {
    const combined = js.venue_address || '';
    const parts = splitAddressLines(combined);
    if (!parts.length) {
      return [js.venue_address1, js.venue_address2, js.venue_address3, [js.venue_town, js.venue_postcode].filter(Boolean).join(' ')].filter(Boolean);
    }
    return parts;
  })();
  setText('VENUE_ADDRESS1', venueLines[0] || '');
  setText('VENUE_ADDRESS2', venueLines[1] || '');
  setText('VENUE_ADDRESS3', venueLines[2] || '');
  setText('VENUE_ADDRESS4', venueLines[3] || [js.venue_town, js.venue_postcode].filter(Boolean).join(' '));
  setText('CATERER_NAME', js.caterer_name || '');
  const displayTotalFees = vatEnabled ? quoteTotal : (totalAmount != null ? totalAmount : quoteTotal);
  setText('TOTAL_FEES', currencyOrEmpty(displayTotalFees), TextAlignment.Right);
  setText('DEPOSIT_AMOUNT', currencyOrEmpty(displayDepositAmount), TextAlignment.Right);
  setText('BALANCE_AMOUNT', currencyOrEmpty(displayBalanceAmount), TextAlignment.Right);
  setText('AHMEN_FEE', currencyOrEmpty(ahmenFee), TextAlignment.Right);
  setText('PRODUCTION_FEES', currencyOrEmpty(productionFees), TextAlignment.Right);
  // Legacy typo fallback (template field PRODUCTION_)FEES)
  setText('PRODUCTION_)FEES', currencyOrEmpty(productionFees), TextAlignment.Right);
  // Common subtotal/VAT/total fields (used by Booking Schedule and others if present)
  setText('SUBTOTAL', currencyOrEmpty(displaySubtotal), TextAlignment.Right);
  setText('VAT', vatEnabled ? currencyOrEmpty(displayVat) : 'N/A', TextAlignment.Right);
  setText('VAT_RATE', vatEnabled ? '20%' : 'N/A', TextAlignment.Right);
  setText('TOTAL', currencyOrEmpty(displayTotal), TextAlignment.Right);
  setText('EXTRA_FEES', currencyOrEmpty(extraFees));
  setText('ADD_PROD', prodLine1);
  setText('ADD_PROD2', prodLine2);
  if (reservedInvoiceNumber != null) {
    setText('INV_NUMBER', `INV-${reservedInvoiceNumber}`);
  }
  setText('BALANCE_DATE', balanceDateDisplay);
  setText('E Special Conditions if any 1', specialLines[0] || '');
  setText('E Special Conditions if any 2', specialLines[1] || '');
  setText('E Special Conditions if any 3', specialLines[2] || '');
  const todayDisplay = formatDisplayDate(new Date());
  const dateAlign = (docTypeRaw === 'quote' || docTypeRaw === 'invoice') ? TextAlignment.Left : TextAlignment.Right;
  setText('TODAY_DATE', todayDisplay, dateAlign);
  setText('DATE_TODAY', todayDisplay, dateAlign);
  setText('DATE_SIGNED', '');

  // Quote-specific fields (static PDF with form fields)
  if (definition?.doc_type === 'quote' || (payload.doc_type || '').toLowerCase() === 'quote' || definitionKey === 'quote') {
    setText('CLIENT_NAME', js.client_name || client.name || '');
    setText('CLIENT_EMAIL', js.client_email || client.email || '');
    setText('CLIENT_PHONE', js.client_phone || client.phone || '');
    setText('CLIENT_ADDRESS1', clientLines[0] || '');
    setText('CLIENT_ADDRESS2', clientLines[1] || '');
    setText('CLIENT_ADDRESS3', clientLines[2] || '');
    setText('CLIENT_ADDRESS4', clientLines[3] || '');
    // DATE_TODAY handled above with quote-specific alignment
    setText('EVENT_DATE', eventDateDisplay);
    setText('VENUE_NAME', js.venue_name || '');
    setText('AHMEN_FEE', currencyOrEmpty(ahmenFee), TextAlignment.Right);
    setText('PRODUCTION_FEES', currencyOrEmpty(productionFees), TextAlignment.Right);
    setText('DEPOSIT_AMOUNT', currencyOrEmpty(displayDepositAmount), TextAlignment.Right);
    setText('BALANCE_AMOUNT', currencyOrEmpty(displayBalanceAmount), TextAlignment.Right);
    setText('SUBTOTAL', currencyOrEmpty(quoteSubtotal), TextAlignment.Right);
    const vatDisplay = vatEnabled ? currencyOrEmpty(effectiveVat) : 'N/A';
    ['VAT', 'VAT_AMOUNT', 'VAT_AMT', 'VAT_VALUE'].forEach(name => setText(name, vatDisplay, TextAlignment.Right));
    setText('VAT_RATE', vatEnabled ? '20%' : 'N/A', TextAlignment.Right);
    setText('TOTAL', currencyOrEmpty(quoteTotal), TextAlignment.Right);
  }

  const hasProduction = Boolean((parseAmount(productionFees) || 0) > 0 || prodLine1 || prodLine2);
  setCheck('AhMen to provide SoundAV Please tick if appropriate', hasProduction);
  ['Tick', 'Tick_2', 'Tick_3', 'Tick_4', 'Tick_5'].forEach(name => setCheck(name, true));

  // For quote PDFs, preserve the template's own field appearances; otherwise rebuild with the selected font
  if (docTypeRaw === 'quote') {
    try {
      form.updateFieldAppearances(baseFont);
    } catch (_) {}
  } else {
    try {
      form.updateFieldAppearances(baseFont);
    } catch (_) {}
  }

  const pdfBytes = await pdfDoc.save();
  await fs.promises.writeFile(targetPath, pdfBytes);

  const docType = docTypeRaw;
  const documentDate = payload.document_date || new Date().toISOString();

  let inserted = null;
  if (docTypeRaw === 'invoice' && reservedInvoiceId != null) {
    const reminderDate = invoiceVariant === 'balance' ? (js.balance_reminder_date || null) : null;
    await db.updateDocumentStatus(reservedInvoiceId, {
      file_path: targetPath,
      status: 'issued',
      total_amount: Number.isFinite(invoiceTotal) ? invoiceTotal : 0,
      balance_due: Number.isFinite(invoiceTotal) ? invoiceTotal : 0,
      due_date: payload.balance_due_date ?? payload.due_date ?? null,
      reminder_date: reminderDate
    });
    inserted = { id: reservedInvoiceId, number: reservedInvoiceNumber };
  } else {
    inserted = await db.addDocument({
      business_id: businessId,
      jobsheet_id: payload.jobsheet_id || null,
      doc_type: docType,
      status: 'generated',
      total_amount: totalAmount ?? 0,
      balance_due: balanceAmount ?? totalAmount ?? 0,
      due_date: payload.balance_due_date ?? payload.due_date ?? null,
      file_path: targetPath,
      client_name: js.client_name || client.name || null,
      event_name: js.event_name || js.event_type || null,
      event_date: js.event_date || event.event_date || null,
      document_date: documentDate,
      definition_key: definitionKey,
      invoice_variant: payload.invoice_variant || null
    });
  }

  return {
    ok: true,
    file_path: targetPath,
    document_id: inserted?.id || null,
    number: inserted?.number ?? null,
    additional_outputs: []
  };
}

async function createDocument(payload = {}) {
  const templatePath = payload.template_path || payload.templatePath || '';
  if (templatePath && /\.pdf$/i.test(String(templatePath))) {
    return createPdfDocument(payload);
  }
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
  if (rawType === 'invoice') {
    await syncFinderInvoiceCounter(businessId);
  }
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

  // Deduplicate by exact file_path (prefer definition_key and non-pdf_export rows)
  try {
    const preferDoc = (left, right) => {
      const leftHasDef = Boolean(left?.definition_key);
      const rightHasDef = Boolean(right?.definition_key);
      if (leftHasDef !== rightHasDef) return leftHasDef ? left : right;

      const leftType = String(left?.doc_type || '').toLowerCase();
      const rightType = String(right?.doc_type || '').toLowerCase();
      const leftIsExport = leftType === 'pdf_export';
      const rightIsExport = rightType === 'pdf_export';
      if (leftIsExport !== rightIsExport) return leftIsExport ? right : left;

      const leftId = Number(left?.document_id);
      const rightId = Number(right?.document_id);
      if (Number.isFinite(leftId) && Number.isFinite(rightId)) {
        return leftId >= rightId ? left : right;
      }
      if (Number.isFinite(leftId)) return left;
      if (Number.isFinite(rightId)) return right;
      return left;
    };

    const byPath = new Map();
    for (const doc of enriched) {
      const fp = doc?.file_path;
      if (!fp) continue;
      const existing = byPath.get(fp);
      if (!existing) { byPath.set(fp, doc); continue; }
      byPath.set(fp, preferDoc(existing, doc));
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
            await syncFinderInvoiceCounter(businessId);
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
  let updated = 0;

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
          if (Number.isInteger(jobsheetId)) {
            const existingJobId = existing?.jobsheet_id != null ? Number(existing.jobsheet_id) : null;
            const canRelink = existingJobId !== jobsheetId && isSubPath(resolvedDir, absolutePath);
            if (existingJobId == null || existingJobId === jobsheetId || canRelink) {
              const patch = {};
              if ((existingJobId == null || canRelink) && Number.isInteger(jobsheetId)) patch.jobsheet_id = jobsheetId;
              if (!existing.client_name && snapshot.client_name) patch.client_name = snapshot.client_name;
              if (!existing.event_name && snapshot.event_type) patch.event_name = snapshot.event_type;
              if (!existing.event_date && snapshot.event_date) patch.event_date = snapshot.event_date;
              if (Object.keys(patch).length) {
                await db.updateDocumentStatus(existing.document_id, patch);
                updated += 1;
                results.push({ document_id: existing.document_id, file_path: absolutePath, updated: true });
              }
            }
          }
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

  return { added, updated, records: results };
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
  const termsPdf = '';
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
  const files = [assets.schedule_pdf, assets.deposit_pdf].filter(Boolean);
  if (!files.length) throw new Error('No booking pack PDFs found in the job folder.');
  const to = (snapshot.client_email || '').trim();
  if (!to) throw new Error('Client email missing on jobsheet');
  const subject = `Booking pack – ${(snapshot.client_name || 'Client')} – ${formatDisplayDate(snapshot.event_date)}`;
  const firstName = (snapshot.client_name || '').trim().split(/\s+/)[0] || 'there';
  const body = `Hi ${firstName},\n\nAttached are your booking schedule and deposit invoice. The deposit is payable on contract signing.\n\nThanks,\nMotti`;
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

function isBalanceInvoiceDocument(doc) {
  if (!doc) return false;
  const docType = String(doc.doc_type || '').toLowerCase();
  const defKey = String(doc.definition_key || '').toLowerCase();
  const variant = String(doc.invoice_variant || doc.definition_invoice_variant || '').toLowerCase();
  if (defKey === 'invoice_balance') return true;
  if (docType === 'invoice' && variant === 'balance') return true;
  const name = String(doc.file_path || doc.definition_label || doc.label || '').toLowerCase();
  return docType === 'invoice' && /balance/.test(name);
}

function getDocumentSortDate(doc) {
  if (!doc) return 0;
  const raw = doc.document_date || doc.created_at || doc.updated_at || '';
  const date = raw ? new Date(raw) : null;
  if (date && !Number.isNaN(date.valueOf())) return date.valueOf();
  const id = Number(doc.document_id);
  return Number.isFinite(id) ? id : 0;
}

function pickLatestDocument(list = []) {
  if (!Array.isArray(list) || !list.length) return null;
  return list.reduce((best, current) => {
    if (!best) return current;
    return getDocumentSortDate(current) >= getDocumentSortDate(best) ? current : best;
  }, null);
}

async function listPlannerItems(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id);
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');
  const includeCompleted = options.includeCompleted === true;
  const sync = options.sync !== false;
  const skipFileScan = options.skipFileScan === true || options.skip_file_scan === true;
  const horizonMonthsRaw = options.horizonMonths ?? options.horizon_months;
  const horizonMonths = Number.isFinite(Number(horizonMonthsRaw)) ? Math.max(0, Number(horizonMonthsRaw)) : 2;
  const todayKey = new Date().toISOString().slice(0, 10);
  const horizonDate = horizonMonths > 0 ? addMonthsISO(todayKey, horizonMonths) : todayKey;

  const [jobsheets, documents, existingActions, scheduledEmails] = await Promise.all([
    db.getAhmenJobsheets({ businessId, includeArchived: false }),
    db.getDocuments({ businessId }),
    db.listPlannerActions({ business_id: businessId }),
    db.listScheduledEmails({ business_id: businessId, status: 'pending', limit: 1000 })
  ]);

  const actionMap = new Map();
  (existingActions || []).forEach(action => {
    const dateKey = normalizeDateKey(action.scheduled_for);
    if (!dateKey) return;
    const key = `${action.jobsheet_id}|${action.action_key}|${dateKey}`;
    actionMap.set(key, action);
  });

  const docsByJobsheet = new Map();
  (documents || []).forEach(doc => {
    const id = Number(doc.jobsheet_id);
    if (!Number.isInteger(id)) return;
    const list = docsByJobsheet.get(id) || [];
    list.push(doc);
    docsByJobsheet.set(id, list);
  });

  const scheduledByJobsheet = new Map();
  (scheduledEmails || []).forEach(entry => {
    const id = Number(entry.jobsheet_id);
    if (!Number.isInteger(id)) return;
    const sendAt = normalizeDateKey(entry.send_at || entry.scheduled_for || '');
    if (!sendAt) return;
    const current = scheduledByJobsheet.get(id);
    if (!current || sendAt < current) {
      scheduledByJobsheet.set(id, sendAt);
    }
  });

  const completedStatuses = new Set(['sent', 'done', 'completed', 'dismissed']);
  const items = [];

  for (const js of jobsheets || []) {
    const jobsheetId = Number(js?.jobsheet_id);
    if (!Number.isInteger(jobsheetId)) continue;
    const statusKey = String(js?.status || '').toLowerCase();
    if (statusKey === 'completed') continue;
    const eventDate = normalizeDateKey(js.event_date);
    if (!eventDate) continue;

    const balanceDue = normalizeDateKey(js.balance_due_date) || addDaysISO(eventDate, -10);
    const balanceSend = normalizeDateKey(js.balance_reminder_date) || addDaysISO(eventDate, -20);
    const paymentCheck = balanceDue ? addDaysISO(balanceDue, 1) : '';

    const jobDocs = docsByJobsheet.get(jobsheetId) || [];
    const balanceDoc = pickLatestDocument(jobDocs.filter(isBalanceInvoiceDocument));
    let fallbackInvoicePath = '';
    let fallbackInvoiceNumber = null;
    if (!balanceDoc || !balanceDoc.file_path) {
      try {
        if (!skipFileScan) {
          const files = await module.exports.listJobFolderFiles({
            businessId,
            jobsheetId,
            extensionPattern: '\\.(pdf)$'
          });
          const balanceMatch = (files || []).find(file => (
            /balance/i.test(file?.name || '')
            || /invoice[_\s-]*balance/i.test(file?.name || '')
            || /bal[-_\s]*inv/i.test(file?.name || '')
          ));
          if (balanceMatch?.path) {
            fallbackInvoicePath = balanceMatch.path;
            const name = String(balanceMatch.name || '');
            const numMatch = name.match(/inv[\-\s]?(\d+)/i);
            if (numMatch && numMatch[1]) {
              const parsed = Number(numMatch[1]);
              if (Number.isFinite(parsed)) fallbackInvoiceNumber = parsed;
            }
          }
        }
      } catch (_) {}
    }
    const isPaid = balanceDoc ? (String(balanceDoc.status || '').toLowerCase() === 'paid' || !!balanceDoc.paid_at) : false;

    const buildAction = async (action_key, scheduled_for, extra = {}) => {
      if (!scheduled_for) return null;
      const dateKey = normalizeDateKey(scheduled_for);
      if (!dateKey) return null;
      if (horizonDate && dateKey > horizonDate) return null;
      const mapKey = `${jobsheetId}|${action_key}|${dateKey}`;
      const existing = actionMap.get(mapKey) || null;
      if (!existing && sync) {
        try {
          await db.upsertPlannerAction({
            business_id: businessId,
            jobsheet_id: jobsheetId,
            action_key,
            scheduled_for: dateKey,
            status: 'pending'
          });
        } catch (_err) {}
      }
      const status = existing?.status || 'pending';
      if (!includeCompleted && completedStatuses.has(String(status).toLowerCase())) return null;
      return {
        action_id: existing?.action_id || null,
        business_id: businessId,
        jobsheet_id: jobsheetId,
        action_key,
        scheduled_for: dateKey,
        status,
        last_notified_at: existing?.last_notified_at || null,
        last_email_at: existing?.last_email_at || null,
        last_error: existing?.last_error || null,
        ...extra
      };
    };

    const clientEmail = String(js.client_email || '').trim();
    const balanceAmount = Number(js.balance_amount || 0);
    const balanceItem = balanceDoc || {};
    const invoiceNumber = balanceItem?.number ?? fallbackInvoiceNumber ?? null;
    const invoicePath = balanceItem?.file_path || fallbackInvoicePath || '';
    const scheduledEmailAt = scheduledByJobsheet.get(jobsheetId) || '';
    const hasScheduledEmail = !!scheduledEmailAt;
    const common = {
      client_name: js.client_name || '',
      client_email: clientEmail,
      event_type: js.event_type || '',
      event_date: js.event_date || '',
      balance_due_date: balanceDue || '',
      balance_reminder_date: balanceSend || '',
      balance_amount: Number.isFinite(balanceAmount) ? balanceAmount : 0,
      invoice_number: invoiceNumber,
      invoice_path: invoicePath,
      paid: isPaid,
      scheduled_email_at: scheduledEmailAt
    };

    const balanceSendAction = await buildAction('balance_send', balanceSend, {
      ...common,
      can_send: !!clientEmail && !!invoicePath && !isPaid && !hasScheduledEmail,
      needs_email: !clientEmail && !hasScheduledEmail,
      needs_invoice: !invoicePath && !hasScheduledEmail,
      auto_send: true
    });
    if (balanceSendAction) items.push(balanceSendAction);

    if (!isPaid && balanceDue) {
      const balanceDueAction = await buildAction('balance_due', balanceDue, {
        ...common,
        auto_send: false,
        info_only: true
      });
      if (balanceDueAction) items.push(balanceDueAction);
    }

    if (!isPaid && paymentCheck) {
      const paymentCheckAction = await buildAction('payment_check', paymentCheck, {
        ...common,
        auto_send: false,
        requires_approval: true,
        can_send: !!clientEmail,
        needs_email: !clientEmail
      });
      if (paymentCheckAction) items.push(paymentCheckAction);
    }

    if (isPaid) {
      const paidDate = normalizeDateKey(balanceItem?.paid_at) || balanceDue || eventDate;
      const thankAction = await buildAction('thank_you', paidDate, {
        ...common,
        auto_send: false,
        requires_approval: true,
        can_send: !!clientEmail,
        needs_email: !clientEmail
      });
      if (thankAction) items.push(thankAction);
    }
  }

  items.sort((a, b) => {
    if (a.scheduled_for === b.scheduled_for) {
      return (a.client_name || '').localeCompare(b.client_name || '');
    }
    return String(a.scheduled_for || '').localeCompare(String(b.scheduled_for || ''));
  });

  return { items };
}

async function sendPlannerEmail(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id);
  const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
  const actionKeyRaw = String(options.action_key || options.actionKey || '').toLowerCase();
  if (!Number.isInteger(businessId) || !Number.isInteger(jobsheetId)) {
    throw new Error('businessId and jobsheetId are required');
  }
  if (!actionKeyRaw) throw new Error('actionKey is required');

  const jobsheet = await db.getAhmenJobsheet(jobsheetId);
  if (!jobsheet) throw new Error('Jobsheet not found');
  const to = String(jobsheet.client_email || '').trim();
  if (!to) throw new Error('Client email missing on jobsheet');

  const templateKey = actionKeyRaw === 'payment_check'
    ? 'payment_reminder'
    : (actionKeyRaw === 'thank_you' ? 'thank_you' : 'invoice_balance');

  const [templates, defaults, signatureResult] = await Promise.all([
    module.exports.getMailTemplates({ businessId }),
    module.exports.getDefaultMailTemplates({ businessId }),
    module.exports.getMailSignature({ businessId })
  ]);

  const def = (defaults && defaults[templateKey]) || {};
  const custom = (templates && templates[templateKey]) || {};
  const tpl = { ...def, ...custom };
  const tokens = buildMailTokenMap(jobsheet || {});
  const subject = renderMailTemplate(tpl.subject || '', tokens);
  const bodyRaw = renderMailTemplate(tpl.body || '', tokens);
  const signature = signatureResult?.signature || '';
  const bodyWithSignature = appendSignatureHtml(bodyRaw, signature);
  const body = normalizeEmailHtmlPreserveSignature(bodyWithSignature, {
    baseFamily: 'Arial, Helvetica, sans-serif',
    baseSize: '10pt',
    baseLineHeight: '1.45'
  });

  let attachments = [];
  if (templateKey === 'invoice_balance' || templateKey === 'payment_reminder') {
    const resolved = await resolveTemplateDefaultAttachments({ businessId, jobsheetId, templateKey: 'invoice_balance' });
    attachments = Array.isArray(resolved?.attachments) ? resolved.attachments : [];
    if (templateKey === 'invoice_balance' && !attachments.length) {
      throw new Error('Balance invoice PDF not found. Generate it before sending.');
    }
  }

  await sendMailViaGraph({
    to,
    subject,
    body,
    is_html: true,
    attachments,
    business_id: businessId,
    jobsheet_id: jobsheetId
  });

  return { ok: true };
}

async function sendInternalNotice(options = {}) {
  const subject = String(options.subject || '').trim();
  const body = String(options.body || '').trim();
  if (!subject && !body) return { ok: true };
  const businessId = options.businessId ?? options.business_id ?? null;
  const jobsheetId = options.jobsheetId ?? options.jobsheet_id ?? null;
  await sendMailViaGraph({
    to: 'motti@ahmen.co.uk',
    subject: subject || 'AhMen reminder',
    body: body || '',
    is_html: true,
    business_id: businessId,
    jobsheet_id: jobsheetId,
    skipLog: true
  });
  return { ok: true };
}

async function updatePlannerAction(options = {}) {
  const actionId = options.action_id ?? options.actionId;
  const businessId = Number(options.businessId ?? options.business_id);
  const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
  const status = options.status || null;
  const completedAt = options.completed_at ?? options.completedAt ?? null;
  const lastNotifiedAt = options.last_notified_at ?? options.lastNotifiedAt ?? null;
  const lastEmailAt = options.last_email_at ?? options.lastEmailAt ?? null;
  const lastError = options.last_error ?? options.lastError ?? null;

  if (actionId != null) {
    const result = await db.updatePlannerActionById({
      action_id: actionId,
      status,
      completed_at: completedAt,
      last_notified_at: lastNotifiedAt,
      last_email_at: lastEmailAt,
      last_error: lastError
    });
    broadcastJobsheetChange({
      type: 'planner-updated',
      businessId: Number.isInteger(businessId) ? businessId : null,
      jobsheetId: Number.isInteger(jobsheetId) ? jobsheetId : null
    });
    return result;
  }

  const actionKey = String(options.action_key || options.actionKey || '').trim();
  const scheduledFor = normalizeDateKey(options.scheduled_for || options.scheduledFor || '');
  if (!Number.isInteger(businessId) || !Number.isInteger(jobsheetId) || !actionKey || !scheduledFor) {
    throw new Error('action_id or full planner key is required');
  }
  const result = await db.upsertPlannerAction({
    business_id: businessId,
    jobsheet_id: jobsheetId,
    action_key: actionKey,
    scheduled_for: scheduledFor,
    status,
    completed_at: completedAt,
    last_notified_at: lastNotifiedAt,
    last_email_at: lastEmailAt,
    last_error: lastError
  });
  broadcastJobsheetChange({
    type: 'planner-updated',
    businessId,
    jobsheetId
  });
  return result;
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
        attachments.push(bundle?.schedule_pdf || '', bundle?.deposit_pdf || '');
      } catch (_) {}
      if (!attachments.filter(Boolean).length) {
        const schedule = findInFolder([
          /schedule/i,
          /booking\s*schedule/i,
          /itinerary/i
        ]);
        const deposit = findInFolder([
          /deposit/i,
          /\bdep\b/i,
          /invoice[-_\s]*deposit/i
        ]);
        attachments.push(schedule || '', deposit || '');
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
    // Enforce page setup to keep layout to one page wide (prevents horizontal overflow across pages)
    try {
      workbook.eachSheet(ws => {
        try {
          ws.pageSetup = Object.assign({}, ws.pageSetup || {}, {
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0, // unlimited height (multiple pages tall ok)
            paperSize: 9 // A4
          });
        } catch (_) {}
      });
    } catch (_) {}

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
    if (rawType === 'invoice') {
      await syncFinderInvoiceCounter(businessId);
    }
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

    // If a manual invoice number is provided, purge phantom DB rows with the same number (no file on disk)
    if (options.invoice_number != null && Number.isFinite(Number(options.invoice_number))) {
      try {
        const dupRows = await db.getDocumentsByNumber(businessId, 'invoice', Number(options.invoice_number));
        if (Array.isArray(dupRows) && dupRows.length) {
          for (const r of dupRows) {
            const p = r && r.file_path ? String(r.file_path) : '';
            let exists = false;
            try { if (p) { await fs.promises.access(p, fs.constants.F_OK); exists = true; } } catch (_) { exists = false; }
            if (!exists) {
              try { await db.deleteDocument(r.document_id); } catch (_) {}
            }
          }
        }
      } catch (_) {}
    }

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

    // Use flat main folder; defer final naming until number known
    const directory = path.resolve(business.save_path);
    await ensureDirectoryExists(directory);
    // Temporary workbook path to avoid collisions; will rename after writing
    let workbookPath = path.join(directory, `.tmp_mcms_invoice_${Date.now()}_${Math.random().toString(36).slice(2)}.xlsx`);

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
    const mcmsKeys = [
      'client_name','client_email','client_phone','client_address1','client_address2','client_town','client_postcode',
      'invoice_date','due_date','total_amount','invoice_code'
    ];
    const fieldKeySet = new Set(['invoice_code', ...mcmsKeys, ...Array.from(placeholderKeys)]);
    const valueSources = await db.getMergeFieldValueSources(Array.from(fieldKeySet)) || {};
    // Inject a contextPath source for invoice_code -> context.invoiceCode
    valueSources['invoice_code'] = { source_type: 'contextPath', source_path: 'invoiceCode' };
    // Ensure MCMS-relevant placeholders point to context paths (override AhMen defaults)
    const ensure = (key, path) => { if (!valueSources[key]) valueSources[key] = { source_type: 'contextPath', source_path: path }; };
    ensure('client_name', 'client.name');
    ensure('client_email', 'client.email');
    ensure('client_phone', 'client.phone');
    ensure('client_address1', 'client.address1');
    ensure('client_address2', 'client.address2');
    ensure('client_town', 'client.town');
    ensure('client_postcode', 'client.postcode');
    // Map invoice_date to context.issueDate; keep legacy issue_date mapping for compatibility
    ensure('invoice_date', 'issueDate');
    ensure('issue_date', 'issueDate');
    ensure('due_date', 'dueDate');
    ensure('total_amount', 'totalAmount');

    // Apply explicit field overrides passed from UI to match template tokens exactly
    const overrides = (options.field_values || options.placeholders || {});
    if (overrides && typeof overrides === 'object') {
      for (const [key, value] of Object.entries(overrides)) {
        if (!key) continue;
        const exactKey = String(key);
        let v = value;
        if (/date/i.test(exactKey)) {
          const dt = toExcelDate(value);
          v = dt || value;
        } else if (typeof value === 'string' && /amount|total|rate|qty|quantity|hours/i.test(exactKey)) {
          const cleaned = String(value).replace(/[^0-9.\-]+/g, '');
          const n = Number(cleaned);
          v = Number.isFinite(n) ? n : value;
        }
        valueSources[exactKey] = { source_type: 'literal', literal_value: v };
        // Make sure replacement can resolve exact and slug forms to this key
        placeholderMap.set(exactKey.toLowerCase(), exactKey);
        placeholderMap.set(exactKey.replace(/[^a-z0-9]+/gi, '_'), exactKey);
      }
    }

    // Helper: find a specific token cell (e.g., invoice_code) prior to replacement
    const findTokenCellAddress = (tokenName) => {
      const m = String(tokenName || '').toLowerCase();
      let address = '';
      workbook.eachSheet(ws => {
        if (address) return;
        ws.eachRow(row => {
          row.eachCell(cell => {
            if (address) return;
            const v = cell && cell.value;
            const scan = (text) => {
              const re = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g; let mm; re.lastIndex = 0;
              while ((mm = re.exec(text)) !== null) {
                const key = (mm[1] || '').toLowerCase();
                if (key === m) { address = cell.address; break; }
              }
            };
            if (typeof v === 'string') scan(v);
            else if (v && typeof v === 'object' && Array.isArray(v.richText)) {
              const text = v.richText.map(f => f && f.text ? f.text : '').join('');
              scan(text);
            }
          });
        });
      });
      return address;
    };

    // Helper: classify item field name from token
    const classifyItemField = (name) => {
      const s = String(name || '').toLowerCase();
      if (/(desc|description|name)$/.test(s)) return 'description';
      if (/(qty|quantity|hours)$/.test(s)) return 'quantity';
      if (/unit$/.test(s)) return 'unit';
      if (/(rate|price|cost)$/.test(s)) return 'rate';
      if (/(amount|line[_-]?total|total)$/.test(s)) return 'amount';
      if (/(date)$/.test(s)) return 'date';
      return '';
    };

    // Helper: write repeatable item rows based on a template row (or block) containing item tokens
    const writeRepeatableItemRows = () => {
      let targetSheet = null;
      let templateRowNumber = null;
      let colToField = new Map(); // col -> field
      let baseBlockRowCount = 1; // number of contiguous template rows with item tokens
      const tokenRe = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;
      workbook.eachSheet(ws => {
        if (templateRowNumber != null) return;
        ws.eachRow((row, r) => {
          if (templateRowNumber != null) return;
          const mapping = new Map();
          row.eachCell((cell, c) => {
            const v = cell && cell.value;
            const check = (text) => {
              let m; tokenRe.lastIndex = 0;
              while ((m = tokenRe.exec(text)) !== null) {
                const raw = (m[1] || '').toLowerCase();
                const norm = raw.replace(/[^a-z0-9]+/g, '_');
                if (norm.startsWith('item')) {
                  const field = classifyItemField(norm.replace(/^item[_\.]?/, ''));
                  if (field) mapping.set(c, field);
                }
              }
            };
            if (typeof v === 'string') check(v);
            else if (v && typeof v === 'object' && Array.isArray(v.richText)) {
              const t = v.richText.map(f => f && f.text ? f.text : '').join('');
              check(t);
            }
          });
          if (mapping.size > 0) {
            targetSheet = ws; templateRowNumber = r; colToField = mapping;
            // Probe subsequent contiguous rows to count how many template item rows exist
            let rr = r + 1;
            while (true) {
              const nextRow = ws.getRow(rr);
              if (!nextRow) break;
              let hasItemToken = false;
              nextRow.eachCell((cell) => {
                const v = cell && cell.value;
                const scan = (text) => {
                  let m; tokenRe.lastIndex = 0;
                  while ((m = tokenRe.exec(text)) !== null) {
                    const raw = (m[1] || '').toLowerCase();
                    const norm = raw.replace(/[^a-z0-9]+/g, '_');
                    if (norm.startsWith('item')) { hasItemToken = true; break; }
                  }
                };
                if (typeof v === 'string') scan(v);
                else if (v && typeof v === 'object' && Array.isArray(v.richText)) scan(v.richText.map(f => f && f.text ? f.text : '').join(''));
              });
              if (hasItemToken) { baseBlockRowCount += 1; rr += 1; }
              else break;
            }
          }
        });
      });
      if (!targetSheet || templateRowNumber == null) return false;

      // If the template already contains multiple item rows, use them first.
      // Insert extra rows for additional items beyond the base block (duplicate the last row for style)
      if (items.length > baseBlockRowCount) {
        const need = items.length - baseBlockRowCount;
        const dupIndex = templateRowNumber + baseBlockRowCount - 1;
        for (let i = 0; i < need; i++) {
          try { targetSheet.duplicateRow(dupIndex, 1, true); } catch (_) {
            try { targetSheet.spliceRows(dupIndex + 1, 0, []); } catch (_) {}
          }
        }
      } else if (items.length < baseBlockRowCount) {
        // Clear remaining template item rows if fewer items provided
        const reReplace = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;
        for (let rr = templateRowNumber + items.length; rr < templateRowNumber + baseBlockRowCount; rr++) {
          const row = targetSheet.getRow(rr);
          row.eachCell((cell) => {
            const v = cell && cell.value;
            const stripItems = (text) => text.replace(reReplace, (_m, keyRaw) => {
              const k = (keyRaw || '').toLowerCase().replace(/[^a-z0-9]+/g, '_');
              if (!k.startsWith('item')) return _m;
              return '';
            });
            if (typeof v === 'string') { cell.value = stripItems(v); }
            else if (v && typeof v === 'object' && Array.isArray(v.richText)) { const text = v.richText.map(f => f && f.text ? f.text : '').join(''); cell.value = stripItems(text); }
          });
        }
      }

      // Fill each item row
      const reReplace = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;
      items.forEach((it, idx) => {
        const r = targetSheet.getRow(templateRowNumber + idx);
        r.eachCell((cell, c) => {
          const v = cell && cell.value;
          const field = colToField.get(c) || '';
          const computeAmount = () => (it.amount != null && Number.isFinite(it.amount))
            ? it.amount
            : ((Number.isFinite(it.quantity) && Number.isFinite(it.rate)) ? (it.quantity * it.rate) : null);
          const replacements = {
            description: it.description || '',
            quantity: (it.quantity != null && Number.isFinite(it.quantity)) ? it.quantity : null,
            unit: it.unit || '',
            rate: (it.rate != null && Number.isFinite(it.rate)) ? it.rate : null,
            amount: computeAmount(),
            date: it.date || it.item_date || null
          };
          const applyFormat = (f, ce) => {
            if (!ce) return;
            try {
              if (f === 'rate' || f === 'amount') applyNumberFormat(ce, { data_type: 'number', format: 'currency' });
              else if (f === 'quantity') applyNumberFormat(ce, { data_type: 'number', format: 'decimal_2' });
              else if (f === 'date') applyNumberFormat(ce, { data_type: 'date' });
              if (f === 'description') {
                ce.alignment = Object.assign({}, ce.alignment || {}, { wrapText: true, vertical: 'top' });
              }
            } catch (_) {}
          };

          // Replace item tokens within the cell text if present
          if (typeof v === 'string') {
            // If this column is a typed field (e.g., date), prefer typed value over inline text replacement
            if (field === 'date') {
              const dt = toExcelDate(replacements.date);
              cell.value = dt || '';
              applyFormat(field, cell);
            } else {
              const newText = v.replace(reReplace, (_m, keyRaw) => {
                const k = (keyRaw || '').toLowerCase().replace(/[^a-z0-9]+/g, '_');
                if (!k.startsWith('item')) return _m; // leave other tokens for later
                const f = classifyItemField(k.replace(/^item[_\.]?/, ''));
                const val = replacements[f];
                if (val == null) return '';
                if (typeof val === 'number' && (f === 'rate' || f === 'amount')) return String(val);
                return String(val);
              });
              cell.value = newText;
              if (field) applyFormat(field, cell);
            }
          } else if (v && typeof v === 'object' && Array.isArray(v.richText)) {
            const text = v.richText.map(f => f && f.text ? f.text : '').join('');
            if (field === 'date') {
              const dt = toExcelDate(replacements.date);
              cell.value = dt || '';
              applyFormat(field, cell);
            } else {
              const newText = text.replace(reReplace, (_m, keyRaw) => {
                const k = (keyRaw || '').toLowerCase().replace(/[^a-z0-9]+/g, '_');
                if (!k.startsWith('item')) return _m;
                const f = classifyItemField(k.replace(/^item[_\.]?/, ''));
                const val = replacements[f];
                if (val == null) return '';
                if (typeof val === 'number' && (f === 'rate' || f === 'amount')) return String(val);
                return String(val);
              });
              cell.value = newText; // flatten to string
              if (field) applyFormat(field, cell);
            }
          } else if (field) {
            // Set direct typed values when cell is an item field
            const val = replacements[field];
            if (val == null) {
              cell.value = '';
            } else {
              if (field === 'date') {
                const dt = toExcelDate(val);
                cell.value = dt || '';
                applyFormat(field, cell);
              } else {
                cell.value = val;
                applyFormat(field, cell);
              }
            }
          }
        });
      });

      return true;
    };

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
    await syncFinderInvoiceCounter(businessId);
    const insert = await db.addDocument({
      business_id: businessId,
      doc_type: 'invoice',
      number: (options.invoice_number != null && Number.isFinite(Number(options.invoice_number))) ? Number(options.invoice_number) : undefined,
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

    // First, write repeatable item rows (if a template row exists)
    let itemRowWritten = false;
    try { if (items.length) itemRowWritten = writeRepeatableItemRows(); } catch (_) { itemRowWritten = false; }
    // Fallback: write items to a simple anchor {{items}} if present
    if (!itemRowWritten && items.length) {
      const writeLineItems = () => {
        let anchor = null;
        let anchorSheet = null;
        workbook.eachSheet(ws => {
          if (anchor) return;
          ws.eachRow((row, rowNumber) => {
            if (anchor) return;
            row.eachCell((cell, colNumber) => {
              if (anchor) return;
              const v = cell && cell.value;
              const isToken = typeof v === 'string' && /^{{\s*items\s*}}$/i.test(v.trim());
              if (isToken) { anchor = { row: rowNumber, col: colNumber }; anchorSheet = ws; }
            });
          });
        });
        if (!anchor || !anchorSheet) return false;
        try { anchorSheet.getCell(anchor.row, anchor.col).value = ''; } catch (_) {}
        const sr = anchorSheet.getRow(anchor.row);
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
          const amountVal = (it.amount != null && Number.isFinite(it.amount)) ? it.amount : ((Number.isFinite(it.quantity) && Number.isFinite(it.rate)) ? (it.quantity * it.rate) : null);
          amtCell.value = amountVal;
          try { if (sr) {
            const c1 = sr.getCell(anchor.col), c2 = sr.getCell(anchor.col+1), c3 = sr.getCell(anchor.col+2), c4 = sr.getCell(anchor.col+3), c5 = sr.getCell(anchor.col+4);
            if (c1 && c1.style) descCell.style = { ...c1.style };
            if (c2 && c2.style) qtyCell.style = { ...c2.style };
            if (c3 && c3.style) unitCell.style = { ...c3.style };
            if (c4 && c4.style) rateCell.style = { ...c4.style };
            if (c5 && c5.style) amtCell.style = { ...c5.style };
          } } catch (_) {}
          try { applyNumberFormat(rateCell, { data_type: 'number', format: 'currency' }); } catch (_) {}
          try { applyNumberFormat(amtCell, { data_type: 'number', format: 'currency' }); } catch (_) {}
        });
        return true;
      };
      try { writeLineItems(); } catch (_) {}
    }

    // Replace remaining placeholders (non-item tokens) after items are populated
    replaceWorkbookPlaceholders(workbook, valueSources, enrichedContext, placeholderMap);
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    sanitizeWorkbookValues(workbook);
    await workbook.xlsx.writeFile(workbookPath);

    // Move and rename to flat main folder with number-first naming: INV-#### - Client - YYYY-MM-DD
    const baseDir = path.resolve(business.save_path);
    await ensureDirectoryExists(baseDir);
    const clientNameSafe = sanitizeFilenameSegment(context.client?.name || '');
    const invoiceDateToken = (options && options.field_values && options.field_values.invoice_date) ? options.field_values.invoice_date : (payload.document_date || new Date().toISOString());
    const invoiceDateIso = formatDateISO(invoiceDateToken);
    const fileBase = [
      `INV-${number}`,
      clientNameSafe || null,
      invoiceDateIso || null
    ].filter(Boolean).map(sanitizeFilenameSegment).join(' - ');

    let newWorkbookPath = path.join(baseDir, `${fileBase}.xlsx`);
    {
      let k = 2;
      while (await pathExists(newWorkbookPath)) {
        newWorkbookPath = path.join(baseDir, `${fileBase} (${k}).xlsx`);
        k += 1;
        if (k > 1000) break;
      }
    }
    try {
      if (!isSubPath(baseDir, workbookPath) || path.basename(workbookPath) !== path.basename(newWorkbookPath)) {
        await fs.promises.rename(workbookPath, newWorkbookPath);
      }
      workbookPath = newWorkbookPath;
    } catch (_) {
      // If rename fails, fall back to using the new path for PDF output
      workbookPath = newWorkbookPath;
    }

    // Build PDF file name and export with stamping (same base as workbook)
    const baseName = path.basename(workbookPath, '.xlsx');
    let pdfPath = path.join(baseDir, `${baseName}.pdf`);
    {
      let n = 2;
      while (await pathExists(pdfPath)) {
        pdfPath = path.join(baseDir, `${baseName} (${n}).pdf`);
        n += 1;
        if (n > 1000) break;
      }
    }

    // Choose stamp cell: prefer cell with {{invoice_code}}, fallback to F14 per template, else E9
    let stampCell = findTokenCellAddress('invoice_code') || 'F14' || 'E9';
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
  // Create a numbered Excel workbook by copying the template verbatim and opening it in Excel.
  // Options: { business_id, definition_key='invoice_balance', invoice_number, client_name, invoice_date }
  createNumberedWorkbookSimple: async (options = {}) => {
    const businessId = Number(options.business_id ?? options.businessId);
    if (!Number.isInteger(businessId)) throw new Error('business_id is required');
    const definitionKey = options.definition_key || 'invoice_balance';
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) throw new Error('Documents folder not configured for this business.');
    const def = await db.getDocumentDefinition(businessId, definitionKey);
    const templatePath = def?.template_path || options.template_path || options.templatePath;
    if (!templatePath) throw new Error('Template path is not configured for this document. Set it in Templates.');
    const resolvedTemplate = path.resolve(templatePath);
    await ensureFileAccessible(resolvedTemplate);

    // Reserve/validate number now
    await syncFinderInvoiceCounter(businessId);
    const insert = await db.addDocument({
      business_id: businessId,
      doc_type: 'invoice',
      number: (options.invoice_number != null && Number.isFinite(Number(options.invoice_number))) ? Number(options.invoice_number) : undefined,
      status: 'draft',
      total_amount: null,
      balance_due: null,
      due_date: null,
      client_name: options.client_name || null,
      document_date: options.invoice_date || new Date().toISOString(),
      definition_key: definitionKey
    });
    const number = insert?.number != null ? Number(insert.number) : null;
    if (number == null) throw new Error('Failed to reserve invoice number');

    const baseDir = path.resolve(business.save_path);
    await ensureDirectoryExists(baseDir);
    const clientSafe = sanitizeFilenameSegment(options.client_name || '');
    const dateIso = formatDateISO(options.invoice_date || new Date().toISOString());
    const baseName = [ `INV-${number}`, clientSafe || null, dateIso || null ].filter(Boolean).join(' - ');
    let destXlsx = path.join(baseDir, `${baseName}.xlsx`);
    {
      let k = 2;
      while (await pathExists(destXlsx)) {
        destXlsx = path.join(baseDir, `${baseName} (${k}).xlsx`);
        k += 1; if (k > 1000) break;
      }
    }

    // Copy template verbatim to preserve all layout/print settings
    await fs.promises.copyFile(resolvedTemplate, destXlsx);

    // Update DB record with workbook path (not PDF)
    try { await db.updateDocumentStatus(insert.id, { file_path: destXlsx, status: 'draft' }); } catch (_) {}

    // Attempt to determine stamp cells (and sheets) for invoice code and client name
    // Prefer cells with {{invoice_code}} and {{client_name}}; fallback invoice to F14
    let invoiceCell = '';
    let clientCell = '';
    let invoiceSheet = '';
    let clientSheet = '';
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(resolvedTemplate);
      const TOKEN = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;
      wb.eachSheet(ws => {
        ws.eachRow((row) => {
          row.eachCell((cell) => {
            const v = cell && cell.value;
            const scan = (text) => {
              let m; TOKEN.lastIndex = 0; while ((m = TOKEN.exec(text)) !== null) {
                const key = (m[1]||'').toLowerCase();
                if (!invoiceCell && key === 'invoice_code') { invoiceCell = cell.address; invoiceSheet = ws.name || ''; }
                if (!clientCell && key === 'client_name') { clientCell = cell.address; clientSheet = ws.name || ''; }
                if (invoiceCell && clientCell) break;
              }
            };
            if (typeof v === 'string') scan(v);
            else if (v && typeof v === 'object' && Array.isArray(v.richText)) scan(v.richText.map(f=>f&&f.text?f.text:'').join(''));
          });
        });
      });
      if (!invoiceCell) invoiceCell = 'F14';
    } catch (_) { invoiceCell = 'F14'; }
    const invoiceText = `INV-${number}`;
    const clientText = (options.client_name || '').toString();

    // Open in Excel for manual editing and stamp the invoice code into the chosen cell (unsaved change)
    try {
      const osaArgs = [
        '-e', 'on run argv',
        '-e', 'if (count of argv) < 1 then error "Missing path"',
        '-e', 'set workbookPosixPath to item 1 of argv',
        '-e', 'set invoiceCell to item 2 of argv',
        '-e', 'set invoiceText to item 3 of argv',
        '-e', 'set invoiceSheet to item 4 of argv',
        '-e', 'set clientCell to item 5 of argv',
        '-e', 'set clientText to item 6 of argv',
        '-e', 'set clientSheet to item 7 of argv',
        '-e', 'set workbookHfs to (POSIX file workbookPosixPath) as text',
        '-e', 'tell application "Microsoft Excel"',
        '-e', 'launch',
        '-e', 'set wb to open workbook workbook file name workbookHfs',
        '-e', 'try',
        '-e', 'set theSheet to active sheet of wb',
        '-e', 'if (invoiceCell is not "" and invoiceText is not "") then',
        '-e', 'set targetSheet to theSheet',
        '-e', 'try',
        '-e', 'if (invoiceSheet is not "") then set targetSheet to worksheet invoiceSheet of wb',
        '-e', 'end try',
        '-e', 'set value of range invoiceCell of targetSheet to invoiceText',
        '-e', 'end if',
        '-e', 'if (clientCell is not "" and clientText is not "") then',
        '-e', 'set targetSheet2 to theSheet',
        '-e', 'try',
        '-e', 'if (clientSheet is not "") then set targetSheet2 to worksheet clientSheet of wb',
        '-e', 'end try',
        '-e', 'set value of range clientCell of targetSheet2 to clientText',
        '-e', 'end if',
        '-e', 'end try',
        '-e', 'activate',
        '-e', 'end tell',
        '-e', 'end run',
        destXlsx,
        (invoiceCell || ''), (invoiceText || ''), (invoiceSheet || ''), (clientCell || ''), (clientText || ''), (clientSheet || '')
      ];
      await new Promise((resolve) => execFile('osascript', osaArgs, { timeout: 20000 }, () => resolve()));
    } catch (_) {}

    return { ok: true, file_path: destXlsx, number, document_id: insert?.id || null };
  },
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
  listPlannerItems,
  sendPlannerEmail,
  updatePlannerAction,
  sendInternalNotice,
  resolveTemplateDefaultAttachments,
  buildGigInfoHtml,
  buildPersonnelLogHtml,
  buildPersonnelLogText,
  // Read an Excel template and return placeholders present ({{token}} occurrences) with positions
  scanTemplatePlaceholders: async (options = {}) => {
    let filePath = options.filePath || options.file_path || '';
    const businessId = Number(options.businessId ?? options.business_id);
    const definitionKey = options.definition_key || options.definitionKey || 'invoice_balance';
    if (!filePath) {
      if (!Number.isInteger(businessId)) throw new Error('businessId or filePath is required');
      const def = await db.getDocumentDefinition(businessId, definitionKey);
      filePath = def?.template_path || '';
    }
    if (!filePath) throw new Error('Template path not found');
    const resolved = path.resolve(filePath);
    await ensureFileAccessible(resolved);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(resolved);
    const TOKEN = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;
    const tokens = [];
    wb.eachSheet(ws => {
      ws.eachRow((row, r) => {
        row.eachCell((cell, c) => {
          const v = cell && cell.value;
          const scan = (text) => { let m; TOKEN.lastIndex = 0; while ((m = TOKEN.exec(text)) !== null) { if (m[1]) tokens.push({ key: m[1], sheet: ws.name, row: r, col: c, address: cell.address }); } };
          if (typeof v === 'string') scan(v);
          else if (v && typeof v === 'object' && Array.isArray(v.richText)) scan(v.richText.map(f => f && f.text ? f.text : '').join(''));
        });
      });
    });
    // Deduplicate by key/address pairs; also produce unique key list
    const byKey = new Map();
    const unique = [];
    const seenAddr = new Set();
    tokens.forEach(t => {
      const addrKey = `${t.sheet}:${t.address}:${t.key}`;
      if (!seenAddr.has(addrKey)) { seenAddr.add(addrKey); unique.push(t); }
      const arr = byKey.get(t.key) || []; arr.push(t); byKey.set(t.key, arr);
    });
    const keys = Array.from(byKey.keys());
    return { tokens: unique, keys };
  },
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
  // Rename the jobsheet folder (if naming changed) and retitle known document filenames
  renameJobsheetArtifacts: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const jobsheetIdRaw = options.jobsheetId ?? options.jobsheet_id;
    const jobsheetId = jobsheetIdRaw != null ? Number(jobsheetIdRaw) : null;
    if (!Number.isInteger(businessId) || !Number.isInteger(jobsheetId)) {
      throw new Error('businessId and jobsheetId are required');
    }
    const business = await db.getBusinessById(businessId);
    if (!business || !business.save_path) {
      throw new Error('Documents folder not configured for this business.');
    }
    const js = await db.getAhmenJobsheet(jobsheetId);
    if (!js) throw new Error('Jobsheet not found');

    // Compute expected new folder path from current snapshot
    const payload = { business_id: businessId, jobsheet_id: jobsheetId, jobsheet_snapshot: js };
    const context = buildContext(payload, business);
    const expectedFolder = buildOutputDirectory(business, context, payload, 'Documents');
    await ensureDirectoryExists(path.resolve(business.save_path));

    // Find current folder candidate from documents DB
    const docsAll = await db.getDocuments({ businessId });
    const docs = (docsAll || []).filter(d => d && Number(d.jobsheet_id) === jobsheetId);
    let currentFolder = '';
    for (const d of docs) {
      const fp = d?.file_path || '';
      if (!fp) continue;
      const dir = path.dirname(fp);
      // Ignore files in the business root
      if (dir && dir !== path.resolve(business.save_path)) {
        currentFolder = dir; break;
      }
    }

    // If we have an existing folder and it differs, move/merge it to the expected location
    if (currentFolder && path.resolve(currentFolder) !== path.resolve(expectedFolder)) {
      await ensureDirectoryExists(expectedFolder);
      let movedViaRename = false;
      try {
        await fs.promises.rename(currentFolder, expectedFolder);
        movedViaRename = true;
      } catch (_) {
        // Fall back: move files one by one
        try {
          const entries = await fs.promises.readdir(currentFolder, { withFileTypes: true });
          for (const e of entries) {
            if (!e.isFile()) continue;
            const src = path.join(currentFolder, e.name);
            let dst = path.join(expectedFolder, e.name);
            let k = 2;
            while (await pathExists(dst)) {
              const base = path.basename(e.name, path.extname(e.name));
              const ext = path.extname(e.name);
              dst = path.join(expectedFolder, `${base} (${k})${ext}`);
              k += 1; if (k > 1000) break;
            }
            try { await fs.promises.rename(src, dst); } catch (_) {}
          }
          // Attempt to remove old folder if empty
          try { await fs.promises.rmdir(currentFolder); } catch (_) {}
        } catch (_) {}
      }

      // Update DB paths for documents under the old folder
      const prefix = path.resolve(currentFolder);
      for (const d of docs) {
        const fp = d?.file_path || '';
        if (!fp) continue;
        const abs = path.resolve(fp);
        if (!abs.startsWith(prefix)) continue;
        const remainder = abs.slice(prefix.length).replace(/^[/\\]+/, '');
        const nextPath = path.join(expectedFolder, remainder);
        try { await db.setDocumentFilePath(d.document_id, nextPath); } catch (_) {}
      }
    }

    // Refresh docs (paths may have changed)
    const docsUpdatedAll = await db.getDocuments({ businessId });
    const docsUpdated = (docsUpdatedAll || []).filter(d => d && Number(d.jobsheet_id) === jobsheetId);

    // Rename known files to match current naming conventions (skip locked)
    const eventDateIso = formatDateISO(js.event_date || '');
    const clientSafe = sanitizeFilenameSegment(js.client_name || '');
    // Workbook (use the actual definition label for this document, not a generic "Workbook" label)
    for (const d of docsUpdated) {
      if ((d.doc_type || '').toLowerCase() !== 'workbook') continue;
      if (d.is_locked) continue;
      const cur = d.file_path || '';
      if (!cur || !cur.toLowerCase().endsWith('.xlsx')) continue;
      let def = null;
      try {
        def = await db.getDocumentDefinition(businessId, d.definition_key || 'workbook');
      } catch (_) {
        def = null;
      }
      const naming = buildFileName(context, payload, def || { label: 'Excel Workbook', key: d.definition_key || 'workbook' });
      const dir = path.resolve(expectedFolder);
      const expectedPath = path.join(dir, naming.fileName);
      const curAbs = path.resolve(cur);
      if (path.resolve(expectedPath) !== curAbs) {
        try {
          await ensureDirectoryExists(dir);
          await fs.promises.rename(curAbs, expectedPath);
          await db.setDocumentFilePath(d.document_id, expectedPath);
        } catch (_) {}
      }
    }

    const renameInvoicesToRoot = options.renameInvoicesToRoot === true;
    if (renameInvoicesToRoot) {
      // Invoice PDFs in business root
      for (const d of docsUpdated) {
        if ((d.doc_type || '').toLowerCase() !== 'invoice') continue;
        if (d.is_locked) continue;
        const num = d.number != null ? Number(d.number) : null;
        if (!Number.isInteger(num)) continue;
        const dir = path.resolve(business.save_path);
        const dateToken = formatDateISO(d.document_date || js.event_date || new Date().toISOString());
        const base = [
          `INV-${num}`,
          clientSafe || null,
          dateToken || null
        ].filter(Boolean).map(sanitizeFilenameSegment).join(' - ');
        const ext = (d.file_path || '').toLowerCase().endsWith('.xlsx') ? '.xlsx' : '.pdf';
        let expectedPath = path.join(dir, `${base}${ext}`);
        let k = 2;
        while (await pathExists(expectedPath)) {
          expectedPath = path.join(dir, `${base} (${k})${ext}`);
          k += 1; if (k > 1000) break;
        }
        const cur = d.file_path || '';
        if (!cur) continue;
        const curAbs = path.resolve(cur);
        if (path.resolve(expectedPath) !== curAbs) {
          try {
            await fs.promises.rename(curAbs, expectedPath);
            await db.setDocumentFilePath(d.document_id, expectedPath);
          } catch (_) {}
        }
      }
    }

    return { ok: true };
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
        // Use plain text with paragraphs so the composer converts to consistent HTML (<p> with margins)
        body: 'Hi {{ client_first_name|there }},\n\nAttached are your booking schedule and deposit invoice.\n\nPlease review and let me know if anything needs updating.\n\nThanks,\n',
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
    const updateExisting = options.updateExisting !== false;

    // Collect candidate PDFs/XLSX with (INV-###) in their filename
    async function walk(dir, depth = 0, maxDepth = 0) {
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
        if (!/inv[\-\s]?\d+/i.test(e.name)) continue;
        if (lower.endsWith('.pdf') || lower.endsWith('.xlsx')) {
          results.push(full);
        }
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
        if (!dateOk || !client) return null; // only treat as strong match if both are present
        const cand = sheets.find(js => norm(js.event_date) === norm(dateOk) && norm(js.client_name) === norm(client));
        return cand || null;
      } catch (_) { return null; }
    };

    function parseFromFilename(filePath) {
      const base = path.basename(filePath);
      const noExt = base.replace(/\.[^.]+$/, '');
      const normalized = noExt.replace(/[–—]+/g, '-');
      const tokens = normalized.split(/\s*-\s*/).filter(Boolean);

      // Invoice number
      let number = null;
      let numIdx = -1;
      for (let i = 0; i < tokens.length; i++) {
        const m = tokens[i].match(/inv[\-\s]?(\d+)/i);
        if (m && m[1]) { number = Number(m[1]); numIdx = i; break; }
      }

      // Date token (prefer ISO; fallback to '14 Jun 2025'/'14 June 2025')
      let dateIso = null;
      let dateIdx = -1;
      const monthMap = { jan:'01', feb:'02', mar:'03', apr:'04', may:'05', jun:'06', jul:'07', aug:'08', sep:'09', oct:'10', nov:'11', dec:'12' };
      for (let j = tokens.length - 1; j >= 0; j--) {
        const t = tokens[j].trim();
        let m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (m) { dateIso = `${m[1]}-${m[2]}-${m[3]}`; dateIdx = j; break; }
        m = t.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\.?\s+(\d{4})$/);
        if (m) {
          const dd = String(m[1]).padStart(2, '0');
          const mm = monthMap[m[2].toLowerCase().slice(0,3)] || null;
          if (mm) { dateIso = `${m[3]}-${mm}-${dd}`; dateIdx = j; break; }
        }
      }

      // Variant
      let variant = null;
      const lower = normalized.toLowerCase();
      if (lower.includes('deposit')) variant = 'deposit'; else if (lower.includes('balance')) variant = 'balance';

      // Client: tokens between number and date
      let clientName = null;
      if (numIdx >= 0) {
        const start = numIdx + 1;
        const end = dateIdx >= 0 ? dateIdx : tokens.length;
        if (end > start) clientName = tokens.slice(start, end).join(' - ').trim() || null;
      }

      return { number, clientName, dateIso, variant };
    }

    // Helper: read file times (prefer creation time)
    async function getFileTimes(p) {
      try {
        const st = await fs.promises.stat(p);
        const birth = st.birthtime && Number.isFinite(st.birthtimeMs) ? new Date(st.birthtimeMs).toISOString() : null;
        const mod = st.mtime && Number.isFinite(st.mtimeMs) ? new Date(st.mtimeMs).toISOString() : null;
        return { birthIso: birth || mod || new Date().toISOString(), mtimeIso: mod || birth || new Date().toISOString() };
      } catch (_) {
        const now = new Date().toISOString();
        return { birthIso: now, mtimeIso: now };
      }
    }

    // Robust filename parsing
    // Accepts formats like:
    //   INV-707 - Client Name - 2025-10-07.xlsx
    //   INV 707 Client Name (2025-10-07).pdf
    //   INV-707_Client_Name_2025-10-07.pdf
    // Client segment is everything between the number and the trailing date, if present.
    function parseFromFilename(fp) {
      const base = path.basename(fp);
      const noExt = base.replace(/\.[^.]+$/, '');
      const norm = noExt
        .replace(/[–—]+/g, '-')       // em/en dashes -> hyphen
        .replace(/_/g, ' ')            // underscores -> space
        .replace(/\s{2,}/g, ' ')      // collapse spaces
        .trim();

      // 1) Extract invoice number (required)
      const numMatch = norm.match(/\bINV[\-\s]?(\d+)\b/i);
      const number = numMatch ? Number(numMatch[1]) : null;
      const afterNum = numMatch ? norm.slice(numMatch.index + numMatch[0].length).trim() : '';

      // 2) Extract date (ISO YYYY-MM-DD), allow parentheses
      let dateIso = null;
      let dateIdx = -1;
      const parenDate = afterNum.match(/\((\d{4}-\d{2}-\d{2})\)\s*$/);
      if (parenDate && parenDate[1]) {
        dateIso = parenDate[1];
        dateIdx = afterNum.lastIndexOf(parenDate[0]);
      }
      if (!dateIso) {
        const tailIso = afterNum.match(/(\d{4}-\d{2}-\d{2})\s*$/);
        if (tailIso && tailIso[1]) {
          dateIso = tailIso[1];
          dateIdx = afterNum.lastIndexOf(tailIso[1]);
        }
      }

      // 3) Client is the middle segment between number and date (or all remaining if no date)
      let clientRaw = afterNum;
      if (dateIdx >= 0) clientRaw = afterNum.slice(0, dateIdx).trim();
      // Trim leading separators like '-' or '—'
      clientRaw = clientRaw.replace(/^[-–—\s]+/, '').trim();
      // Remove trailing separators if left
      clientRaw = clientRaw.replace(/[-–—\s]+$/, '').trim();
      const clientName = clientRaw || null;

      // 4) Variant via keywords in whole name
      const lowerAll = norm.toLowerCase();
      const variant = lowerAll.includes('deposit') ? 'deposit' : (lowerAll.includes('balance') ? 'balance' : null);

      return { number, clientName, dateIso, variant };
    }

    let imported = 0;
    for (const anyPath of files) {
      try {
        // Skip if invoice already recorded for this exact file
        // eslint-disable-next-line no-await-in-loop
        const existing = await db.getDocumentByFilePath(businessId, anyPath);
        if (existing && String(existing.doc_type || '').toLowerCase() === 'invoice') {
          if (updateExisting) {
            const base = path.basename(anyPath).toLowerCase();
            const isPdf = base.endsWith('.pdf');
            const isXlsx = base.endsWith('.xlsx');
            const parsedCommon = parseFromFilename(anyPath);
            const times = await getFileTimes(anyPath);
            const patch = {};
            if (parsedCommon.clientName && parsedCommon.clientName !== existing.client_name) patch.client_name = parsedCommon.clientName;
            if (parsedCommon.dateIso) {
              patch.event_date = parsedCommon.dateIso;
            }
            if (!existing.document_date) patch.document_date = times.birthIso;
            if (isPdf && String(existing.status || '').toLowerCase() !== 'issued') patch.status = 'issued';
            try { if (Object.keys(patch).length) await db.updateDocumentStatus(existing.document_id, patch); } catch (_) {}
          }
          continue;
        }

        const base = path.basename(anyPath);
        const parsedCommon = parseFromFilename(anyPath);
        const number = Number.isInteger(parsedCommon.number) ? parsedCommon.number : null;
        if (number == null || !Number.isInteger(number)) continue;

        const lowerName = base.toLowerCase();
        const isPdf = lowerName.endsWith('.pdf');
        const isXlsx = lowerName.endsWith('.xlsx');

        if (isPdf) {
          // If an invoice with this number already exists (e.g., workbook row),
          // update the matching row (same base name) to point to the PDF and mark as issued.
          try {
            // eslint-disable-next-line no-await-in-loop
            const dupRows = await db.getDocumentsByNumber(businessId, 'invoice', number);
            if (Array.isArray(dupRows) && dupRows.length) {
              const base = (p)=>{ const s=(p||'').toString(); const name = s.split(/\\\\|\//).pop() || ''; return name.replace(/\.[^.]+$/, ''); };
              const selBase = base(anyPath);
              // Prefer matching workbook row in the same base name
              let pick = dupRows.find(r => base(r?.file_path||'') === selBase && (r?.file_path||'').toLowerCase().endsWith('.xlsx'))
                      || dupRows.find(r => base(r?.file_path||'') === selBase)
                      || dupRows.find(r => (r?.file_path||'').toLowerCase().endsWith('.xlsx'))
                      || null;
              if (pick && pick.document_id != null) {
                // eslint-disable-next-line no-await-in-loop
                await db.updateDocumentStatus(pick.document_id, { file_path: anyPath, status: 'issued' });
                imported += 1;
                continue;
              }
            }
          } catch (_) {}

          const parsed = parsedCommon;
          const times = await getFileTimes(anyPath);
          const variant = parsed.variant || (lowerName.includes('deposit') ? 'deposit' : (lowerName.includes('balance') ? 'balance' : null));
          const js = matchJobsheet(anyPath);
          const payload = {
            business_id: businessId,
            jobsheet_id: js?.jobsheet_id || null,
            doc_type: 'invoice',
            number,
            status: 'issued',
            total_amount: js ? (variant === 'deposit' ? (js?.deposit_amount ?? null) : (js?.balance_amount ?? null)) : null,
            balance_due: js ? (variant === 'balance' ? (js?.balance_amount ?? null) : (js?.deposit_amount ?? js?.balance_amount ?? null)) : null,
            due_date: parsed.dateIso || (js ? (variant === 'balance' ? (js?.balance_due_date ?? null) : (js?.event_date ?? null)) : null),
            file_path: anyPath,
            client_name: parsed.clientName || js?.client_name || null,
            event_name: js?.event_type || null,
            event_date: parsed.dateIso || js?.event_date || null,
            document_date: times.birthIso,
            definition_key: null,
            invoice_variant: variant
          };
          let createdId = null;
          try {
            // eslint-disable-next-line no-await-in-loop
            const inserted = await db.addDocument(payload);
            createdId = inserted?.id || null;
            imported += 1;
          } catch (insErr) {
            // Duplicate number or other conflict — fall back to inserting without a number, then set requested number explicitly
            try {
              const fallbackPayload = { ...payload };
              delete fallbackPayload.number;
              // eslint-disable-next-line no-await-in-loop
              const inserted = await db.addDocument(fallbackPayload);
              createdId = inserted?.id || null;
              if (createdId != null && Number.isInteger(number)) {
                try { await db.setDocumentNumber(createdId, number); } catch (_) {}
              }
              imported += 1;
            } catch (err2) {
              // eslint-disable-next-line no-console
              console.warn('Invoice import skipped', base, err2?.message || err2);
            }
          }
        } else if (isXlsx) {
          // For workbooks, import as draft invoice so they appear even without a PDF
          // Attempt to parse client and date from filename; use creation time when date missing
          const parsedX = parsedCommon;
          const times = await getFileTimes(anyPath);
          const docDate = parsedX.dateIso ? new Date(parsedX.dateIso).toISOString() : times.birthIso;
          const payload = {
            business_id: businessId,
            doc_type: 'invoice',
            number,
            status: 'draft',
            total_amount: null,
            balance_due: null,
            due_date: null,
            file_path: anyPath,
            client_name: parsedX.clientName || null,
            event_date: parsedX.dateIso || null,
            document_date: docDate,
            definition_key: 'invoice_balance'
          };
          try {
            // eslint-disable-next-line no-await-in-loop
            await db.addDocument(payload);
            imported += 1;
          } catch (insErr) {
            // Duplicate number or conflict — insert without number, then assign number explicitly to keep Excel+PDF paired
            try {
              const fallbackPayload = { ...payload };
              delete fallbackPayload.number;
              // eslint-disable-next-line no-await-in-loop
              const inserted = await db.addDocument(fallbackPayload);
              if (inserted && inserted.id != null && Number.isInteger(number)) {
                try { await db.setDocumentNumber(inserted.id, number); } catch (_) {}
              }
              imported += 1;
            } catch (err2) {
              // eslint-disable-next-line no-console
              console.warn('Workbook import skipped', base, err2?.message || err2);
            }
          }
        }
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('Failed to import invoice file', anyPath, err);
      }
    }

    try {
      const maxNum = await db.getMaxInvoiceNumber(businessId);
      const last = Number.isInteger(Number(maxNum)) ? Number(maxNum) : 0;
      await db.setLastInvoiceNumber(businessId, last);
    } catch (_) {}

    return { imported };
  },
  rebuildInvoiceFromFilename: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const documentId = Number(options.documentId ?? options.document_id);
    if (!Number.isInteger(businessId)) throw new Error('businessId is required');
    if (!Number.isInteger(documentId)) throw new Error('documentId is required');

    const doc = await db.getDocumentById(documentId);
    if (!doc) throw new Error('Document not found');
    const filePath = doc.file_path || '';
    if (!filePath) throw new Error('File path missing for document');

    const lowerName = filePath.toLowerCase();
    const variant = lowerName.includes('deposit') ? 'deposit' : (lowerName.includes('balance') ? 'balance' : null);

    // Try to match jobsheet strictly via folder naming
    const matchJobsheet = (fp) => {
      try {
        const dirBase = path.basename(path.dirname(fp));
        const parts = dirBase.split(' - ');
        const dateStr = parts[0] || '';
        const client = parts[1] || '';
        const dateOk = /^\d{4}-\d{2}-\d{2}$/.test(dateStr) ? dateStr : '';
        if (!dateOk || !client) return null;
        return (options.sheets || []).find(js => (js.event_date || '').slice(0,10) === dateOk && (js.client_name || '').toLowerCase() === client.toLowerCase()) || null;
      } catch (_) { return null; }
    };

    // Load jobsheets for matching
    let sheets = [];
    try { sheets = await db.getAhmenJobsheets({ businessId }); } catch (_) { sheets = []; }
    const js = matchJobsheet(filePath);

    // Parse filename tokens
    const parseFromFilename = (fp) => {
      const base = path.basename(fp);
      const noExt = base.replace(/\.[^.]+$/, '');
      const normalized = noExt.replace(/[–—]+/g, '-');
      const tokens = normalized.split(/\s*-\s*/).filter(Boolean);
      let number = null, numIdx = -1;
      for (let i = 0; i < tokens.length; i++) {
        const m = tokens[i].match(/inv[\-\s]?(\d+)/i);
        if (m && m[1]) { number = Number(m[1]); numIdx = i; break; }
      }
      let dateIso = null, dateIdx = -1;
      const monthMap = { jan:'01', feb:'02', mar:'03', apr:'04', may:'05', jun:'06', jul:'07', aug:'08', sep:'09', oct:'10', nov:'11', dec:'12' };
      for (let j = tokens.length - 1; j >= 0; j--) {
        const t = tokens[j].trim();
        let m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (m) { dateIso = `${m[1]}-${m[2]}-${m[3]}`; dateIdx = j; break; }
        m = t.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\.?\s+(\d{4})$/);
        if (m) {
          const dd = String(m[1]).padStart(2, '0');
          const mm = monthMap[m[2].toLowerCase().slice(0,3)] || null;
          if (mm) { dateIso = `${m[3]}-${mm}-${dd}`; dateIdx = j; break; }
        }
      }
      // Client between number and date
      let clientName = null;
      if (numIdx >= 0) {
        const start = numIdx + 1;
        const end = dateIdx >= 0 ? dateIdx : tokens.length;
        if (end > start) clientName = tokens.slice(start, end).join(' - ').trim() || null;
      }
      // Variant from tokens
      let variant = null;
      const lower = normalized.toLowerCase();
      if (lower.includes('deposit')) variant = 'deposit'; else if (lower.includes('balance')) variant = 'balance';
      return { number, clientName, dateIso, variant };
    };
    const parsed = parseFromFilename(filePath);

    // Determine updates
    const updatePayload = {};
    if (parsed.clientName) updatePayload.client_name = parsed.clientName;
    if (parsed.dateIso) updatePayload.event_date = parsed.dateIso;
    if (js) {
      updatePayload.event_name = js.event_type || null;
      updatePayload.total_amount = variant === 'deposit' ? (js.deposit_amount ?? null) : (js.balance_amount ?? null);
      updatePayload.balance_due = variant === 'balance' ? (js.balance_amount ?? null) : (js.deposit_amount ?? js.balance_amount ?? null);
      updatePayload.due_date = variant === 'balance' ? (js.balance_due_date ?? null) : (js.event_date ?? null);
    }
    if (variant) updatePayload.invoice_variant = variant;

    // Preview mode: return proposed changes without applying
    if (options.preview === true) {
      return {
        ok: true,
        preview: true,
        document_id: documentId,
        parsed,
        matched_jobsheet_id: js?.jobsheet_id || null,
        proposed: { number: parsed.number, ...updatePayload }
      };
    }

    // Apply changes
    const type = String(doc.doc_type || '').toLowerCase();
    if (type !== 'invoice') {
      // Promote to invoice first
      const promoteOpts = Number.isInteger(parsed.number) ? { number: parsed.number } : {};
      const res = await db.promotePdfToInvoice(documentId, promoteOpts);
      const targetId = res?.id || documentId;
      await db.updateDocumentStatus(targetId, updatePayload);
      return { ok: true, document_id: targetId, promoted: true };
    }

    // Update number if present
    if (Number.isInteger(parsed.number)) {
      try { await db.setDocumentNumber(documentId, parsed.number); } catch (_) {}
    }
    await db.updateDocumentStatus(documentId, updatePayload);
    return { ok: true, document_id: documentId, promoted: false };
  },
  relinkInvoiceToJobsheet: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const documentId = Number(options.documentId ?? options.document_id);
    const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
    const preview = options.preview === true;
    if (!Number.isInteger(businessId)) throw new Error('businessId is required');
    if (!Number.isInteger(documentId)) throw new Error('documentId is required');
    if (!Number.isInteger(jobsheetId)) throw new Error('jobsheetId is required');

    const doc = await db.getDocumentById(documentId);
    if (!doc) throw new Error('Document not found');
    const js = await db.getAhmenJobsheet(jobsheetId);
    if (!js) throw new Error('Jobsheet not found');

    const lower = (v) => (v == null ? '' : String(v).toLowerCase());
    let variant = (doc && doc.invoice_variant) ? lower(doc.invoice_variant) : '';
    if (!variant) {
      const fp = doc?.file_path || '';
      const ln = fp.toLowerCase();
      if (ln.includes('deposit')) variant = 'deposit';
      else if (ln.includes('balance')) variant = 'balance';
    }

    const updatePayload = {
      jobsheet_id: jobsheetId,
      client_name: js.client_name || null,
      event_name: js.event_type || null,
      event_date: js.event_date || null,
      invoice_variant: variant || null
    };
    if (variant === 'deposit') {
      updatePayload.total_amount = js.deposit_amount != null ? Number(js.deposit_amount) : null;
      updatePayload.balance_due = updatePayload.total_amount;
      updatePayload.due_date = js.event_date || null;
    } else if (variant === 'balance') {
      updatePayload.total_amount = js.balance_amount != null ? Number(js.balance_amount) : null;
      updatePayload.balance_due = updatePayload.total_amount;
      updatePayload.due_date = js.balance_due_date || null;
    }

    if (preview) {
      return {
        ok: true,
        preview: true,
        document_id: documentId,
        jobsheet_id: jobsheetId,
        proposed: updatePayload
      };
    }

    await db.updateDocumentStatus(documentId, updatePayload);
    return { ok: true, document_id: documentId, jobsheet_id: jobsheetId };
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
  },
  linkPdfToDefinition: async (options = {}) => {
    const businessId = Number(options.businessId ?? options.business_id ?? options.id);
    const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
    const definitionKey = String(options.definitionKey ?? options.definition_key ?? '').trim();
    const rawPath = options.filePath ?? options.file_path;
    if (!Number.isInteger(businessId) || !Number.isInteger(jobsheetId)) {
      throw new Error('businessId and jobsheetId are required');
    }
    if (!definitionKey) {
      throw new Error('definitionKey is required');
    }
    if (!rawPath || typeof rawPath !== 'string') {
      throw new Error('filePath is required');
    }

    const resolvedPath = path.resolve(rawPath);
    await ensureFileAccessible(resolvedPath);

    const [definition, jobsheet, allDocs] = await Promise.all([
      db.getDocumentDefinition(businessId, definitionKey),
      db.getAhmenJobsheet(jobsheetId),
      db.getDocuments({ businessId })
    ]);

    const docType = (definition?.doc_type || 'pdf_export').toLowerCase();
    const invoiceVariant = definition?.invoice_variant || null;
    const status = docType === 'invoice' ? 'issued' : 'exported';

    let docDate = new Date().toISOString();
    try {
      const stats = await fs.promises.stat(resolvedPath);
      if (stats?.mtime instanceof Date && !Number.isNaN(stats.mtime.valueOf())) {
        docDate = stats.mtime.toISOString();
      }
    } catch (_) {}

    const invoiceNumberMatch = docType === 'invoice'
      ? String(path.basename(resolvedPath)).match(/inv[\-\s]?(\d+)/i)
      : null;
    const parsedNumber = invoiceNumberMatch && invoiceNumberMatch[1]
      ? Number(invoiceNumberMatch[1])
      : null;

    const existingForDef = (allDocs || []).find(doc => (
      Number(doc?.jobsheet_id) === jobsheetId && doc?.definition_key === definitionKey
    ));

    const patch = {
      file_path: resolvedPath,
      status,
      jobsheet_id: jobsheetId,
      definition_key: definitionKey,
      doc_type: docType,
      invoice_variant: invoiceVariant,
      client_name: jobsheet?.client_name || undefined,
      event_name: jobsheet?.event_type || undefined,
      event_date: jobsheet?.event_date || undefined,
      document_date: docDate
    };

    if (existingForDef?.document_id != null) {
      await db.updateDocumentStatus(existingForDef.document_id, patch);
      if (Number.isInteger(parsedNumber) && docType === 'invoice') {
        try { await db.setDocumentNumber(existingForDef.document_id, parsedNumber); } catch (_) {}
      }
      broadcastJobsheetChange({
        type: 'documents-updated',
        businessId,
        jobsheetId,
        documentIds: [existingForDef.document_id]
      });
      return { ok: true, updated: true, document_id: existingForDef.document_id };
    }

    const existingByPath = await db.getDocumentByFilePath(businessId, resolvedPath);
    if (existingByPath?.document_id != null) {
      await db.updateDocumentStatus(existingByPath.document_id, patch);
      if (Number.isInteger(parsedNumber) && docType === 'invoice') {
        try { await db.setDocumentNumber(existingByPath.document_id, parsedNumber); } catch (_) {}
      }
      broadcastJobsheetChange({
        type: 'documents-updated',
        businessId,
        jobsheetId,
        documentIds: [existingByPath.document_id]
      });
      return { ok: true, updated: true, document_id: existingByPath.document_id };
    }

    const insert = await db.addDocument({
      business_id: businessId,
      jobsheet_id: jobsheetId,
      doc_type: docType,
      status,
      total_amount: null,
      balance_due: null,
      due_date: null,
      file_path: resolvedPath,
      client_name: jobsheet?.client_name || null,
      event_name: jobsheet?.event_type || null,
      event_date: jobsheet?.event_date || null,
      document_date: docDate,
      definition_key: definitionKey,
      invoice_variant: invoiceVariant
    });
    const insertedId = insert?.id || null;
    if (insertedId != null && Number.isInteger(parsedNumber) && docType === 'invoice') {
      try { await db.setDocumentNumber(insertedId, parsedNumber); } catch (_) {}
    }
    broadcastJobsheetChange({
      type: 'documents-updated',
      businessId,
      jobsheetId,
      documentIds: insertedId != null ? [insertedId] : []
    });
    return { ok: true, added: true, document_id: insertedId };
  }
};

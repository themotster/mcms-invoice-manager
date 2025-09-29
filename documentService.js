const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
let chokidar = null;
try { chokidar = require('chokidar'); } catch (_err) { chokidar = null; }
const ExcelJS = require('exceljs');
const db = require('./db');

const INVALID_FILENAME_CHARS = /[\\/:*?"<>|]/g;
const TEMPLATE_BINDING_KEY = 'ahmen_excel';
const PLACEHOLDER_PATTERN = /{{\s*([a-zA-Z0-9_.-]+)\s*}}/g;

function normalizeTokenKey(value) {
  if (!value) return '';
  return String(value)
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
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
  const osaArgs = [
    '-e', 'on run argv',
    '-e', 'if (count of argv) < 2 then error "Missing arguments"',
    '-e', 'set workbookPosixPath to item 1 of argv',
    '-e', 'set targetPdfPosix to item 2 of argv',
    '-e', 'set workbookHfs to (POSIX file workbookPosixPath) as text',
    '-e', 'set targetPdfHfs to (POSIX file targetPdfPosix) as text',
    '-e', 'set pdfAlias to POSIX file targetPdfPosix',
    '-e', 'tell application "Microsoft Excel"',
    '-e', 'launch',
    '-e', 'activate',
    '-e', 'set wb to missing value',
    '-e', 'try',
    '-e', 'set wb to open workbook workbook file name workbookHfs',
    '-e', 'end try',
    '-e', 'if wb is missing value then',
    '-e', 'tell application "Finder" to open file workbookHfs',
    '-e', 'repeat with i from 1 to 50',
    '-e', 'delay 0.1',
    '-e', 'try',
    '-e', 'set wb to active workbook',
    '-e', 'exit repeat',
    '-e', 'end try',
    '-e', 'end repeat',
    '-e', 'end if',
    '-e', 'if wb is missing value then error "Unable to open workbook"',
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
    '-e', 'end run',
    sourcePath,
    targetPath
  ];

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

    try {
      // Treat existing PDFs as immutable: do not overwrite existing files
      try {
        const exists = await pathExists(targetPdfPath);
        if (exists) {
          // If DB has a locked record, report locked; otherwise report already exists
          if (businessId != null) {
            try {
              const existing = await db.getDocumentByFilePath(businessId, targetPdfPath);
              if (existing && existing.is_locked) {
                outputs.push({ success: false, sheet: info.suffix, error: 'PDF is locked' });
                continue;
              }
            } catch (_err) {}
          }
          outputs.push({ success: false, sheet: info.suffix, error: 'PDF already exists' });
          continue;
        }
      } catch (_err) {}

      // If a PDF already exists and is locked in DB, block export (defensive)
      if (businessId != null) {
        try {
          const existing = await db.getDocumentByFilePath(businessId, targetPdfPath);
          if (existing && existing.is_locked) {
            outputs.push({ success: false, sheet: info.suffix, error: 'PDF is locked' });
            continue;
          }
        } catch (_err) {}
      }

      await saveWorkbookAsPdf(workbookPath, targetPdfPath, { activeSheetOnly });
      outputs.push({
        success: true,
        sheet: info.suffix,
        label: `PDF`,
        file_path: targetPdfPath
      });
    } catch (err) {
      outputs.push({
        success: false,
        sheet: info.suffix,
        error: err?.message || 'Unable to export sheet'
      });
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
  return { ok, workbook_path: normalizedPath, outputs, message };
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

  await db.deleteDocument(id);
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
      '-e', 'activate',
      '-e', 'set wb to missing value',
      '-e', 'try',
      '-e', 'set wb to open workbook workbook file name workbookHfs',
      '-e', 'end try',
      '-e', 'if wb is missing value then',
      '-e', 'tell application "Finder" to open file workbookHfs',
      '-e', 'repeat with i from 1 to 50',
      '-e', 'delay 0.1',
      '-e', 'try',
      '-e', 'set wb to active workbook',
      '-e', 'exit repeat',
      '-e', 'end try',
      '-e', 'end repeat',
      '-e', 'end if',
      '-e', 'if wb is missing value then error "Unable to open workbook"',
      '-e', 'try',
      '-e', 'close workbook wb saving no',
      '-e', 'end try',
      '-e', 'end tell',
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
module.exports = {
  normalizeTemplate,
  createDocument,
  exportWorkbookPdfs,
  deleteDocument,
  syncJobsheetOutputs,
  watchDocumentsFolder,
  unwatchDocumentsFolder,
  filterDocumentsByExistingFiles,
  listJobsheetDocuments,
  preflightPdfExport
  ,
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

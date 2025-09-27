const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const db = require('./db');

const INVALID_FILENAME_CHARS = /[\\/:*?"<>|]/g;
const TEMPLATE_BINDING_KEY = 'ahmen_excel';
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
  ahmen_fee: 'context.totalAmount',
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

  if (format === 'date_human') {
    return formatDateHuman(rawValue);
  }

  if (dataType === 'number') {
    const numeric = Number(rawValue);
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

async function ensureUniquePath(directory, fileName) {
  let candidate = path.join(directory, fileName);
  if (!fs.existsSync(candidate)) return candidate;

  const ext = path.extname(fileName);
  const name = path.basename(fileName, ext);
  let counter = 1;
  while (true) {
    const attempt = path.join(directory, `${name} (${counter})${ext}`);
    if (!fs.existsSync(attempt)) return attempt;
    counter += 1;
  }
}

async function createWorkbookDocument(payload = {}) {
  const docType = (payload.doc_type || '').toLowerCase();
  if (docType !== 'workbook') {
    throw new Error('Only the workbook document type is supported in this workflow.');
  }

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

  const targetPath = await ensureUniquePath(directory, naming.fileName);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const bindings = await db.getMergeFieldBindingsByTemplate(TEMPLATE_BINDING_KEY);
  const fieldKeys = [...new Set((bindings || []).map(binding => binding.field_key).filter(Boolean))];
  const valueSources = await db.getMergeFieldValueSources(fieldKeys);

  await fillWorkbook(workbook, bindings, valueSources, context);
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

  const watcher = fs.watch(rootPath, { recursive: true }, () => {
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

module.exports = {
  normalizeTemplate,
  createDocument,
  deleteDocument,
  syncJobsheetOutputs,
  watchDocumentsFolder,
  unwatchDocumentsFolder
};

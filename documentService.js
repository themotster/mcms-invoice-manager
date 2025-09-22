const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const db = require('./db');

const AHMEN_TEMPLATE_KEY = 'ahmen_excel';

function ensureDirectoryExists(dirPath) {
  if (!dirPath) return;
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function sanitizeForFilename(text) {
  return (text || 'Untitled')
    .replace(/[\\/:*?"<>|]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseDate(value) {
  if (!value) return new Date();
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return new Date();
  return parsed;
}

function formatDateParts(dateInput) {
  const date = parseDate(dateInput);
  const iso = date.toISOString().slice(0, 10);
  const human = date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'long',
    year: 'numeric'
  });
  return { iso, human };
}

function escapeXmlText(value) {
  if (value === undefined || value === null) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/\r?\n/g, '&#10;');
}

function normalizeNumericValue(value) {
  if (value === undefined || value === null || value === '') return null;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return null;
  return Math.round(numeric * 100) / 100;
}

function formatNumericForCell(value) {
  const normalized = normalizeNumericValue(value);
  if (normalized === null) return null;
  if (Number.isInteger(normalized)) return String(normalized);
  return normalized.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
}

function ensureCalcPrAttributes(xml) {
  if (!xml) return xml;
  const attrCleanup = attrs => attrs
    .replace(/\sfullCalcOnLoad="[^"]*"/i, '')
    .replace(/\scalcOnSave="[^"]*"/i, '');

  if (/<calcPr[^>]*\/>/i.test(xml)) {
    return xml.replace(/<calcPr([^>]*)\/>/i, (_match, attrs = '') => {
      const nextAttrs = attrCleanup(attrs);
      return `<calcPr${nextAttrs} fullCalcOnLoad="1" calcOnSave="1"/>`;
    });
  }

  if (/<calcPr[^>]*>/i.test(xml)) {
    return xml.replace(/<calcPr([^>]*)>/i, (_match, attrs = '') => {
      const nextAttrs = attrCleanup(attrs);
      return `<calcPr${nextAttrs} fullCalcOnLoad="1" calcOnSave="1">`;
    });
  }

  return xml.replace(/<\/workbook>/i, '<calcPr fullCalcOnLoad="1" calcOnSave="1"/>\n</workbook>');
}

async function removeCalcChainArtifacts(zip) {
  const calcChainPath = 'xl/calcChain.xml';
  if (!zip.file(calcChainPath)) return;

  zip.remove(calcChainPath);

  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const workbookRelsEntry = zip.file(workbookRelsPath);
  if (workbookRelsEntry) {
    let relsXml = await workbookRelsEntry.async('string');
    relsXml = relsXml.replace(/<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain"[^>]*\/>\s*/gi, '');
    zip.file(workbookRelsPath, relsXml);
  }

  const contentTypesPath = '[Content_Types].xml';
  const contentEntry = zip.file(contentTypesPath);
  if (contentEntry) {
    let contentXml = await contentEntry.async('string');
    contentXml = contentXml.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>\s*/gi, '');
    zip.file(contentTypesPath, contentXml);
  }
}

function replaceAhmenSharedStrings(xml, replacements) {
  let nextXml = xml;
  Object.entries(replacements).forEach(([placeholder, rawValue]) => {
    const safeValue = escapeXmlText(rawValue || '');
    const pattern = new RegExp(`>${placeholder}<`, 'g');
    nextXml = nextXml.replace(pattern, `>${safeValue}<`);
  });
  return nextXml;
}

function replaceAhmenStringCell(xml, cellRef, value) {
  if (!cellRef) return xml;
  const rowNumber = cellRef.replace(/[^0-9]/g, '');
  if (!rowNumber) return xml;

  const rowPattern = new RegExp(`<row[^>]*r="${rowNumber}"[\s\S]*?<\/row>`, 'i');
  const rowMatch = xml.match(rowPattern);
  if (!rowMatch) return xml;

  const safeValue = escapeXmlText(value || '');
  const cellPattern = new RegExp(`<c[^>]*r="${cellRef}"[\s\S]*?<\/c>`, 'i');

  let rowXml = rowMatch[0];
  const replacement = safeValue
    ? `<c r="${cellRef}" t="inlineStr"><is><t>${safeValue}</t></is></c>`
    : `<c r="${cellRef}"/>`;

  if (cellPattern.test(rowXml)) {
    rowXml = rowXml.replace(cellPattern, replacement);
  } else {
    rowXml = rowXml.replace('</row>', `${replacement}</row>`);
  }

  return xml.replace(rowPattern, rowXml);
}

function buildNumericCell(cellRef, styleId, value) {
  const numeric = formatNumericForCell(value);
  const styleAttr = styleId ? ` s="${styleId}"` : '';
  if (numeric === null) {
    return `<c r="${cellRef}"${styleAttr}/>`;
  }
  return `<c r="${cellRef}"${styleAttr}><v>${numeric}</v></c>`;
}

function replaceAhmenNumericCell(xml, cellRef, styleId, value) {
  const cellPattern = new RegExp(`<c[^>]*r=\"${cellRef}\"[^>]*>([\\s\\S]*?)<\/c>`, 'i');
  const selfClosingPattern = new RegExp(`<c[^>]*r=\"${cellRef}\"[^>]*/>`, 'i');
  const replacement = buildNumericCell(cellRef, styleId, value);

  const replaced = xml.replace(cellPattern, replacement);
  if (replaced !== xml) {
    return replaced;
  }
  const replacedSelf = xml.replace(selfClosingPattern, replacement);
  if (replacedSelf !== xml) {
    return replacedSelf;
  }
  return xml;
}

function buildSheetPathMap(workbookXml, workbookRelsXml) {
  const sheetMap = {};
  if (!workbookXml || !workbookRelsXml) return sheetMap;

  const relMap = {};
  const relRegex = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/gi;
  let match = relRegex.exec(workbookRelsXml);
  while (match) {
    relMap[match[1]] = match[2];
    match = relRegex.exec(workbookRelsXml);
  }

  const sheetRegex = /<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>/gi;
  match = sheetRegex.exec(workbookXml);
  while (match) {
    const name = match[1];
    const rid = match[2];
    const target = relMap[rid];
    if (name && target) {
      sheetMap[name] = target;
    }
    match = sheetRegex.exec(workbookXml);
  }

  return sheetMap;
}

function normalizeNumericOutput(value) {
  if (value === undefined || value === null || value === '') return null;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return null;
  return numeric;
}

function formatBindingStringValue(rawValue, bindingFormat) {
  if (rawValue === undefined || rawValue === null) return '';
  if (!bindingFormat) return String(rawValue ?? '');

  switch (bindingFormat) {
    case 'date_human': {
      const input = Array.isArray(rawValue) ? rawValue[0] : rawValue;
      if (!input) return '';
      return formatDateParts(input).human;
    }
    default:
      return String(rawValue ?? '');
  }
}

function buildMergeFieldValueMap(context) {
  const map = {};
  const business = context.business || {};
  const client = context.client || {};
  const event = context.event || {};

  map.business_name = business.business_name || '';

  map.client_name = client.name || '';
  map.client_email = client.email || '';
  map.client_phone = client.phone || '';
  map.client_address1 = client.address1 || client.address || '';
  map.client_address2 = client.address2 || '';
  map.client_address3 = client.address3 || '';
  map.client_town = client.town || '';
  map.client_postcode = client.postcode || '';

  map.event_type = event.event_name || event.type || '';
  map.event_date = event.event_date || context.documentDate || '';
  map.event_start = event.startTime || event.event_start || '';
  map.event_end = event.endTime || event.event_end || '';

  map.venue_name = event.venue_name || '';
  map.venue_address1 = event.venue_address1 || '';
  map.venue_address2 = event.venue_address2 || '';
  map.venue_address3 = event.venue_address3 || '';
  map.venue_town = event.venue_town || event.town || '';
  map.venue_postcode = event.venue_postcode || event.postcode || '';
  map.caterer_name = event.caterer_name || event.catererName || '';

  map.total_amount = context.totalAmount;
  map.extra_fees = context.extraFees;
  map.production_fees = context.productionFees;
  map.deposit_amount = context.depositAmount ?? context.deposit;
  map.balance_amount = context.balanceAmount;

  map.balance_due_date = context.balanceDate || context.dueDate || '';
  map.balance_reminder_date = context.balanceRemind || '';

  map.service_types = context.serviceType || '';
  map.specialist_singers = context.specialistSingers || '';

  return map;
}

async function applyAhmenTemplateWithZip({ templatePath, destinationPath, context }) {
  if (!templatePath || !fs.existsSync(templatePath)) {
    throw new Error('AhMen template is missing');
  }

  const templateBuffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);
  const sharedStringsPath = 'xl/sharedStrings.xml';
  const workbookPath = 'xl/workbook.xml';
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';

  const sharedStringsEntry = zip.file(sharedStringsPath);
  const workbookEntry = zip.file(workbookPath);
  const workbookRelsEntry = zip.file(workbookRelsPath);
  if (!sharedStringsEntry || !workbookEntry || !workbookRelsEntry) {
    throw new Error('AhMen template is missing required workbook data');
  }

  let sharedStringsXml = await sharedStringsEntry.async('string');
  let workbookXml = await workbookEntry.async('string');
  const workbookRelsXml = await workbookRelsEntry.async('string');

  const sheetPathMap = buildSheetPathMap(workbookXml, workbookRelsXml);
  const sheetCache = new Map();

  const loadSheetRecord = async (sheetName) => {
    if (sheetCache.has(sheetName)) return sheetCache.get(sheetName);
    const relativePath = sheetPathMap[sheetName];
    if (!relativePath) return null;
    const entry = zip.file(`xl/${relativePath}`);
    if (!entry) return null;
    const xml = await entry.async('string');
    const record = { path: relativePath, xml };
    sheetCache.set(sheetName, record);
    return record;
  };

  const bindings = await db.getMergeFieldBindingsByTemplate(AHMEN_TEMPLATE_KEY);
  const fieldValues = buildMergeFieldValueMap(context);
  const sharedReplacements = {};

  for (const binding of bindings) {
    const rawValue = fieldValues[binding.field_key];

    if (binding.placeholder) {
      const token = `{${binding.placeholder}}`;
      if (!(token in sharedReplacements)) {
        const formatted = binding.data_type === 'number'
          ? (normalizeNumericOutput(rawValue) ?? '')
          : formatBindingStringValue(rawValue, binding.format);
        sharedReplacements[token] = formatted;
      }
    }

    if (!binding.sheet || !binding.cell) continue;
    const sheetRecord = await loadSheetRecord(binding.sheet);
    if (!sheetRecord) continue;

    if (binding.data_type === 'number') {
      const numericValue = normalizeNumericOutput(rawValue);
      sheetRecord.xml = replaceAhmenNumericCell(sheetRecord.xml, binding.cell, binding.style || '12', numericValue);
    } else {
      const stringValue = formatBindingStringValue(rawValue, binding.format);
      sheetRecord.xml = replaceAhmenStringCell(sheetRecord.xml, binding.cell, stringValue);
    }
  }

  if (Object.keys(sharedReplacements).length) {
    sharedStringsXml = replaceAhmenSharedStrings(sharedStringsXml, sharedReplacements);
  }

  zip.file(sharedStringsPath, sharedStringsXml);

  sheetCache.forEach(record => {
    zip.file(`xl/${record.path}`, record.xml);
  });

  if (workbookXml) {
    workbookXml = ensureCalcPrAttributes(workbookXml);
    zip.file(workbookPath, workbookXml);
  }

  await removeCalcChainArtifacts(zip);

  const output = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(destinationPath, output);
}

function defaultExtensionForType(docType) {
  switch ((docType || '').toLowerCase()) {
    case 'contract':
      return '.docx';
    case 'gig_sheet':
      return '.xlsx';
    case 'quote':
    case 'invoice':
      return '.xlsx';
    default:
      return '.txt';
  }
}

function pickTemplatePath(business, docType) {
  if (!business || !docType) return null;
  const key = docType.toLowerCase();
  if (key === 'invoice') return business.invoice_template_path || null;
  if (key === 'quote') return business.quote_template_path || null;
  if (key === 'contract') return business.contract_template_path || null;
  if (key === 'gig_sheet') return business.gig_sheet_template_path || null;
  return null;
}

function buildDestinationPath({ business, client, event, docType, number, templatePath }) {
  const templateExists = templatePath && fs.existsSync(templatePath);
  const extension = templateExists
    ? (path.extname(templatePath) || defaultExtensionForType(docType))
    : '.txt';

  const dateSource = event?.event_date || new Date().toISOString();
  const { iso, human } = formatDateParts(dateSource);
  const clientName = sanitizeForFilename(client?.name || business?.business_name || 'Client');
  const humanFormatted = sanitizeForFilename(human);

  const fileNameParts = [iso, clientName, humanFormatted].filter(Boolean);
  const fileName = `${fileNameParts.join(' - ')}${extension}`;
  const baseDir = business?.save_path || path.join(process.cwd(), 'documents');
  ensureDirectoryExists(baseDir);

  return path.join(baseDir, fileName);
}

function replacePlaceholdersInWorksheet(worksheet, replacements) {
  worksheet.eachRow(row => {
    row.eachCell(cell => {
      if (typeof cell.value === 'string') {
        let nextValue = cell.value;
        Object.entries(replacements).forEach(([placeholder, value]) => {
          if (!placeholder) return;
          if (nextValue.includes(placeholder)) {
            nextValue = nextValue.replace(new RegExp(placeholder, 'g'), value ?? '');
          }
        });
        cell.value = nextValue;
      }
    });
    row.commit?.();
  });
}

function setCellIfExists(worksheet, address, value) {
  if (!worksheet || !address) return;
  const cell = worksheet.getCell(address);
  if (cell) {
    if (value === undefined || value === null || value === '') {
      return;
    }
    cell.value = value;
  }
}

function findRowWithPlaceholder(worksheet, placeholder) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= row.cellCount; colNumber += 1) {
      const cell = row.getCell(colNumber);
      if (typeof cell.value === 'string' && cell.value.includes(placeholder)) {
        return rowNumber;
      }
    }
  }
  return null;
}

function fillLineItemsWithTemplateRow(worksheet, templateRowNumber, lineItems, columnMapping) {
  const items = Array.isArray(lineItems) ? lineItems : [];
  const row = worksheet.getRow(templateRowNumber);

  if (!items.length) {
    Object.values(columnMapping).forEach(col => {
      const cell = row.getCell(col);
      if (typeof cell.value === 'string' && cell.value.includes('{')) {
        cell.value = '';
      }
    });
    row.commit();
    return;
  }

  if (items.length > 1) {
    worksheet.duplicateRow(templateRowNumber, items.length - 1, true);
  }

  items.forEach((item, index) => {
    const targetRowNumber = templateRowNumber + index;
    const targetRow = worksheet.getRow(targetRowNumber);
    const dateValue = item.date ? formatDateParts(item.date).human : '';
    if (columnMapping.date) targetRow.getCell(columnMapping.date).value = dateValue;
    if (columnMapping.description) targetRow.getCell(columnMapping.description).value = item.description || '';
    if (columnMapping.duration) targetRow.getCell(columnMapping.duration).value = item.duration || '';
    if (columnMapping.amount) {
      const amountValue = item.amount != null ? Number(item.amount) : '';
      targetRow.getCell(columnMapping.amount).value = amountValue;
    }
    targetRow.commit();
  });
}

function buildCommonReplacements({
  business,
  client,
  event,
  docType,
  number,
  documentDate,
  dueDate,
  totalAmount,
  balanceDue
}) {
  const documentDateParts = formatDateParts(documentDate);
  const dueDateParts = dueDate ? formatDateParts(dueDate) : { iso: '', human: '' };
  return {
    '{BUSINESS_NAME}': business?.business_name || '',
    '{CLIENT_NAME}': client?.name || '',
    '{CLIENT_EMAIL}': client?.email || '',
    '{CLIENT_PHONE}': client?.phone || '',
    '{CLIENT_CONTACT}': client?.contact || '',
    '{CLIENT_ADDRESS1}': client?.address1 || client?.address || '',
    '{CLIENT_ADDRESS2}': client?.address2 || '',
    '{CLIENT_ADDRESS3}': client?.address || '',
    '{CLIENT_TOWN}': client?.town || '',
    '{CLIENT_POSTCODE}': client?.postcode || '',
    '{EVENT_NAME}': event?.event_name || '',
    '{EVENT_TYPE}': event?.event_name || '',
    '{EVENT_DATE}': event?.event_date ? formatDateParts(event.event_date).human : '',
    '{VENUE_NAME}': event?.venue_name || '',
    '{VENUE_ADDRESS1}': event?.venue_address1 || '',
    '{VENUE_ADDRESS2}': event?.venue_address2 || '',
    '{VENUE_ADDRESS3}': '',
    '{VENUE_TOWN}': event?.town || '',
    '{VENUE_POSTCODE}': event?.postcode || '',
    '{DOCUMENT_NUMBER}': number != null ? String(number) : '',
    '{NUMBER}': number != null ? String(number) : '',
    '{INVOICE_NUMBER}': number != null ? String(number) : '',
    '{QUOTE_NUMBER}': number != null ? String(number) : '',
    '{INVOICE_DATE}': documentDateParts.human,
    '{QUOTE_DATE}': documentDateParts.human,
    '{DOCUMENT_DATE}': documentDateParts.human,
    '{DUE_DATE}': dueDateParts.human,
    '{TOTAL}': totalAmount != null ? Number(totalAmount) : '',
    '{BALANCE_DUE}': balanceDue != null ? Number(balanceDue) : ''
  };
}

function fillMcmsInvoiceWorksheet(worksheet, context) {
  const replacements = buildCommonReplacements(context);
  replacePlaceholdersInWorksheet(worksheet, replacements);

  // Header fields
  setCellIfExists(worksheet, 'F14', context.number);
  setCellIfExists(worksheet, 'D14', replacements['{INVOICE_DATE}']);
  setCellIfExists(worksheet, 'C15', context.client?.name || '');

  // Line items start at row 18
  fillLineItemsWithTemplateRow(worksheet, 18, context.lineItems, {
    date: 'B',
    description: 'C',
    duration: 'E',
    amount: 'F'
  });

  // Clear any unused rows up to 15 lines
  for (let j = context.lineItems.length; j < 15; j++) {
    const row = 18 + j;
    worksheet.getCell(`B${row}`).value = "";
    worksheet.getCell(`C${row}`).value = "";
    worksheet.getCell(`E${row}`).value = "";
    worksheet.getCell(`F${row}`).value = "";
  }

  // Totals
  if (context.totalAmount != null) {
    setCellIfExists(worksheet, 'F40', context.totalAmount);
  }
  if (context.balanceDue != null) {
    setCellIfExists(worksheet, 'F41', context.balanceDue);
  }
}

async function generateExcelDocument({ templatePath, destinationPath, business, docType, context }) {
  if (business?.business_name === 'AhMen A Cappella Ltd') {
    await applyAhmenTemplateWithZip({ templatePath, destinationPath, context });
    return;
  }

  const workbook = new ExcelJS.Workbook();
  if (templatePath && fs.existsSync(templatePath)) {
    const extension = path.extname(templatePath).toLowerCase();
    if (extension === '.xlsm') {
      await workbook.xlsx.readFile(templatePath);
    } else {
      await workbook.xlsx.readFile(templatePath);
    }
  } else {
    workbook.addWorksheet('Sheet1');
  }

  if (workbook.calcProperties) {
    workbook.calcProperties.fullCalcOnLoad = true;
    workbook.calcProperties.calcOnSave = true;
  }
  if (workbook.model && workbook.model.calcChain) {
    delete workbook.model.calcChain;
  }

  if (business?.business_name === 'Motti Cohen Music Services' && docType === 'invoice') {
    const sheet = workbook.worksheets[0];
    if (sheet) {
      fillMcmsInvoiceWorksheet(sheet, context);
    }
  } else {
    workbook.worksheets.forEach(sheet => {
      const replacements = buildCommonReplacements(context);
      replacePlaceholdersInWorksheet(sheet, replacements);
    });
  }

  await workbook.xlsx.writeFile(destinationPath);
}

function writePlaceholderFile(destinationPath, context) {
  const lines = [
    `Document Type: ${context.docType || ''}`,
    `Number: ${context.number || ''}`,
    `Client: ${context.client?.name || ''}`,
    `Event: ${context.event?.event_name || ''}`,
    `Generated at: ${new Date().toISOString()}`
  ];
  fs.writeFileSync(destinationPath, lines.join('\n'), 'utf8');
}

function copyTemplate(templatePath, destinationPath) {
  fs.copyFileSync(templatePath, destinationPath);
}

async function generateDocumentFile({ templatePath, destinationPath, business, docType, context }) {
  try {
    const extension = path.extname(destinationPath).toLowerCase();
    if ((extension === '.xlsx' || extension === '.xlsm')) {
      await generateExcelDocument({ templatePath, destinationPath, business, docType, context });
    } else if (templatePath && fs.existsSync(templatePath)) {
      copyTemplate(templatePath, destinationPath);
    } else {
      writePlaceholderFile(destinationPath, context);
    }
  } catch (err) {
    throw new Error(`Unable to generate document file: ${err.message}`);
  }
}

function moveFileToTrash(filePath) {
  if (!filePath) return null;
  try {
    if (!fs.existsSync(filePath)) {
      return null;
    }
    const originalDir = path.dirname(filePath);
    const trashDir = path.join(originalDir, '.trash');
    ensureDirectoryExists(trashDir);

    const base = path.basename(filePath);
    let targetPath = path.join(trashDir, base);
    let counter = 1;
    while (fs.existsSync(targetPath)) {
      targetPath = path.join(trashDir, `${base}.${counter}`);
      counter += 1;
    }
    fs.renameSync(filePath, targetPath);
    return targetPath;
  } catch (err) {
    console.error('Failed to move file to trash:', err);
    throw err;
  }
}

function findAvailablePath(targetPath) {
  if (!targetPath) return targetPath;
  if (!fs.existsSync(targetPath)) return targetPath;

  const directory = path.dirname(targetPath);
  const extension = path.extname(targetPath);
  const baseName = path.basename(targetPath, extension);

  let counter = 1;
  let candidate = path.join(directory, `${baseName} (${counter})${extension}`);
  while (fs.existsSync(candidate)) {
    counter += 1;
    candidate = path.join(directory, `${baseName} (${counter})${extension}`);
  }
  return candidate;
}

function moveFile(source, destination) {
  try {
    fs.renameSync(source, destination);
  } catch (err) {
    if (err?.code === 'EXDEV') {
      fs.copyFileSync(source, destination);
      fs.unlinkSync(source);
      return;
    }
    throw err;
  }
}

function appendFileSuffix(filePath, suffix) {
  if (!filePath || !suffix) return filePath;
  const extension = path.extname(filePath);
  const base = extension ? filePath.slice(0, -extension.length) : filePath;
  return `${base}${suffix}${extension}`;
}

async function relocateBusinessDocuments({ businessId, sourcePath, targetPath }) {
  if (!businessId) {
    throw new Error('businessId is required to relocate documents');
  }
  if (!targetPath) {
    throw new Error('targetPath is required to relocate documents');
  }

  const normalizedTarget = path.resolve(targetPath);
  const normalizedSource = sourcePath ? path.resolve(sourcePath) : null;
  ensureDirectoryExists(normalizedTarget);

  const documents = await db.getDocuments({ businessId });
  const summary = {
    moved: [],
    skipped: [],
    errors: []
  };

  for (const doc of documents) {
    const documentId = doc?.document_id;
    const currentPath = doc?.file_path;
    if (!documentId || !currentPath) {
      summary.skipped.push({ documentId, reason: 'missingPath' });
      continue;
    }

    const absoluteCurrent = path.resolve(currentPath);
    if (!fs.existsSync(absoluteCurrent)) {
      summary.skipped.push({ documentId, reason: 'missingFile', path: absoluteCurrent });
      continue;
    }

    let relativePath = path.basename(absoluteCurrent);
    if (normalizedSource) {
      const candidate = path.relative(normalizedSource, absoluteCurrent);
      if (candidate.startsWith('..')) {
        summary.skipped.push({ documentId, reason: 'outsideSource', path: absoluteCurrent });
        continue;
      }
      const hasUnsafeSegment = candidate.split(path.sep).some(segment => segment === '..');
      if (hasUnsafeSegment) {
        summary.skipped.push({ documentId, reason: 'unsafeRelativePath', path: absoluteCurrent });
        continue;
      }
      relativePath = candidate || path.basename(absoluteCurrent);
    }

    const destinationPath = path.join(normalizedTarget, relativePath);
    ensureDirectoryExists(path.dirname(destinationPath));
    const finalPath = findAvailablePath(destinationPath);

    try {
      moveFile(absoluteCurrent, finalPath);
      await db.updateDocumentStatus(documentId, { file_path: finalPath });
      summary.moved.push({ documentId, from: absoluteCurrent, to: finalPath });
    } catch (err) {
      summary.errors.push({ documentId, error: err?.message || String(err) });
    }
  }

  return summary;
}

async function createDocument(documentData) {
  if (!documentData?.doc_type) {
    throw new Error('Document type is required');
  }
  if (!documentData?.business_id) {
    throw new Error('business_id is required to create a document');
  }

  const rawJobsheetId = documentData?.jobsheet_id;
  const jobsheetId = rawJobsheetId != null ? Number(rawJobsheetId) : null;
  const normalizedJobsheetId = Number.isInteger(jobsheetId) ? jobsheetId : null;

  const business = await db.getBusinessById(documentData.business_id);
  if (!business) {
    throw new Error('Business not found for document creation');
  }

  let event = null;
  let client = null;

  if (documentData.event_id) {
    event = await db.getEventById(documentData.event_id);
    if (event?.client_id) {
      client = await db.getClientById(event.client_id);
    }
  }

  if (!client && documentData.client_id) {
    client = await db.getClientById(documentData.client_id);
  }

  const manualFilePath = documentData.file_path ? path.resolve(documentData.file_path) : null;
  const documentDate = documentData.document_date || new Date().toISOString();
  const insertPayload = { ...documentData };
  delete insertPayload.file_path;
  const sessionIds = Array.isArray(documentData.session_ids) ? documentData.session_ids.slice() : [];
  delete insertPayload.session_ids;
  const lineItems = Array.isArray(documentData.line_items) ? documentData.line_items.slice() : [];
  const clientOverride = documentData.client_override ? { ...documentData.client_override } : null;
  const eventOverride = documentData.event_override ? { ...documentData.event_override } : null;
  const paymentTerms = documentData.payment_terms || documentData.payment_terms_text || '';
  const notes = documentData.notes || '';
  const quoteMeta = documentData.quote_meta ? { ...documentData.quote_meta } : {};
  const contractMeta = documentData.contract_meta ? { ...documentData.contract_meta } : {};
  const discountAmount = documentData.discount_amount;
  const depositAmount = documentData.deposit_amount;
  const extraFees = documentData.extra_fees;
  const productionFees = documentData.production_fees;
  const serviceType = documentData.service_types;
  const specialistSingers = documentData.specialist_singers;
  const balanceAmountOverride = documentData.balance_amount;
  const balanceDateOverride = documentData.balance_due_date;
  const balanceRemindOverride = documentData.balance_reminder_date;
  const invoiceVariant = documentData.invoice_variant || null;
  const fileNameSuffix = documentData.file_name_suffix || '';
  const footerText = documentData.footer || business?.document_footer || '';

  delete insertPayload.line_items;
  delete insertPayload.client_override;
  delete insertPayload.event_override;
  delete insertPayload.payment_terms;
  delete insertPayload.payment_terms_text;
  delete insertPayload.notes;
  delete insertPayload.quote_meta;
  delete insertPayload.contract_meta;
  delete insertPayload.discount_amount;
  delete insertPayload.deposit_amount;
  delete insertPayload.extra_fees;
  delete insertPayload.production_fees;
  delete insertPayload.service_types;
  delete insertPayload.specialist_singers;
  delete insertPayload.invoice_variant;
  delete insertPayload.file_name_suffix;
  delete insertPayload.footer;

  const resolvedClientName = (clientOverride?.name || client?.name || documentData.client_name || '').trim();
  const resolvedEventName = (eventOverride?.event_name || eventOverride?.type || event?.event_name || documentData.event_name || '').trim();
  const resolvedEventDate = eventOverride?.event_date || event?.event_date || documentData.event_date || null;

  insertPayload.client_name = resolvedClientName || null;
  insertPayload.event_name = resolvedEventName || null;
  insertPayload.event_date = resolvedEventDate || null;
  insertPayload.document_date = documentDate;
  insertPayload.jobsheet_id = normalizedJobsheetId;

  const calculatedTotal = lineItems.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
  if ((insertPayload.total_amount == null || insertPayload.total_amount === '') && calculatedTotal) {
    insertPayload.total_amount = calculatedTotal;
  }
  if (insertPayload.balance_due == null || insertPayload.balance_due === '') {
    insertPayload.balance_due = insertPayload.total_amount ?? calculatedTotal;
  }

  if (!insertPayload.due_date && documentData.doc_type === 'invoice' && documentData.document_date) {
    insertPayload.due_date = documentData.document_date;
  }

  const insertResult = await db.addDocument(insertPayload);

  let templatePath = pickTemplatePath(business, documentData.doc_type);
  if (!templatePath && business?.business_name === 'AhMen A Cappella Ltd') {
    templatePath = path.resolve(__dirname, 'AhMen Client Data and Docs Template.xlsx');
  }

  const clientForPath = clientOverride
    ? { ...(client || {}), ...clientOverride, name: resolvedClientName || clientOverride.name }
    : (resolvedClientName ? { ...(client || {}), name: resolvedClientName } : client);
  const eventForPath = eventOverride
    ? { ...(event || {}), ...eventOverride, event_name: resolvedEventName || eventOverride.event_name, event_date: resolvedEventDate || eventOverride.event_date }
    : (resolvedEventName || resolvedEventDate
      ? { ...(event || {}), event_name: resolvedEventName || event?.event_name, event_date: resolvedEventDate || event?.event_date }
      : event);

  let destinationPath = manualFilePath || buildDestinationPath({
    business,
    client: clientForPath,
    event: eventForPath,
    docType: documentData.doc_type,
    number: insertResult.number,
    templatePath
  });

  if (manualFilePath) {
    ensureDirectoryExists(path.dirname(manualFilePath));
  }

  const suffixLabel = fileNameSuffix
    || (invoiceVariant === 'deposit' ? ' - Deposit' : invoiceVariant === 'balance' ? ' - Balance' : '');
  if (suffixLabel) {
    destinationPath = appendFileSuffix(destinationPath, suffixLabel);
  }

  const context = {
    business,
    client: clientOverride ? { ...client, ...clientOverride } : client,
    event: eventOverride ? { ...event, ...eventOverride } : event,
    docType: documentData.doc_type,
    number: insertResult.number,
    documentDate,
    dueDate: insertPayload.due_date,
    totalAmount: insertPayload.total_amount,
    balanceDue: insertPayload.balance_due,
    lineItems,
    paymentTerms,
    notes,
    quoteMeta,
    contractMeta,
    discountAmount,
    depositAmount,
    deposit: depositAmount,
    extraFees,
    productionFees,
    serviceType,
    specialistSingers,
    balanceAmount: balanceAmountOverride ?? insertPayload.balance_due,
    balanceDate: balanceDateOverride ?? insertPayload.due_date,
    balanceRemind: balanceRemindOverride ?? '',
    invoiceVariant,
    footer: footerText
  };

  await generateDocumentFile({
    templatePath,
    destinationPath,
    business,
    docType: documentData.doc_type,
    context
  });

  await db.updateDocumentStatus(insertResult.id, { file_path: destinationPath });
  await Promise.all(sessionIds.map(sessionId => db.markSessionExported(sessionId, true)));

  return {
    id: insertResult.id,
    number: insertResult.number,
    file_path: destinationPath,
    jobsheet_id: normalizedJobsheetId
  };
}

async function deleteDocument(documentId, options = {}) {
  const removeFile = !!options.removeFile;
  const document = await db.getDocumentById(documentId);
  if (!document) {
    throw new Error('Document not found');
  }

  let trashedPath = null;
  if (removeFile && document.file_path) {
    try {
      trashedPath = moveFileToTrash(document.file_path);
    } catch (err) {
      throw new Error(`Unable to remove document file: ${err.message}`);
    }
  }

  await db.deleteDocument(documentId);

  return {
    trashedPath,
    removedFile: removeFile && !!trashedPath
  };
}

async function normalizeTemplateFile(templatePath) {
  const fallbackPath = path.resolve(__dirname, 'AhMen Client Data and Docs Template.xlsx');
  const targetPath = templatePath ? path.resolve(templatePath) : fallbackPath;

  if (!fs.existsSync(targetPath)) {
    throw new Error(`Template not found at ${targetPath}`);
  }

  const buffer = fs.readFileSync(targetPath);
  const zip = await JSZip.loadAsync(buffer);

  const calcChainPath = 'xl/calcChain.xml';
  if (zip.file(calcChainPath)) {
    zip.remove(calcChainPath);
  }

  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const relsEntry = zip.file(workbookRelsPath);
  if (relsEntry) {
    let relsXml = await relsEntry.async('string');
    relsXml = relsXml.replace(/<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain"[^>]*\/>\s*/gi, '');
    zip.file(workbookRelsPath, relsXml);
  }

  const contentTypesPath = '[Content_Types].xml';
  const contentEntry = zip.file(contentTypesPath);
  if (contentEntry) {
    let contentXml = await contentEntry.async('string');
    contentXml = contentXml.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>\s*/gi, '');
    zip.file(contentTypesPath, contentXml);
  }

  const workbookPath = 'xl/workbook.xml';
  const workbookEntry = zip.file(workbookPath);
  if (workbookEntry) {
    let workbookXml = await workbookEntry.async('string');
    const cleanup = attrs => attrs
      .replace(/\sfullCalcOnLoad="[^"]*"/i, '')
      .replace(/\scalcOnSave="[^"]*"/i, '');

    const selfClosingPattern = /<calcPr([^>]*)\/>/i;
    const openPattern = /<calcPr([^>]*)>/i;

    if (selfClosingPattern.test(workbookXml)) {
      workbookXml = workbookXml.replace(selfClosingPattern, (_match, attrs = '') => {
        const nextAttrs = cleanup(attrs || '');
        return `<calcPr${nextAttrs} fullCalcOnLoad="1" calcOnSave="1"/>`;
      });
    } else if (openPattern.test(workbookXml)) {
      workbookXml = workbookXml.replace(openPattern, (_match, attrs = '') => {
        const nextAttrs = cleanup(attrs || '');
        return `<calcPr${nextAttrs} fullCalcOnLoad="1" calcOnSave="1">`;
      });
    } else {
      workbookXml = workbookXml.replace(/<\/workbook>/i, '<calcPr fullCalcOnLoad="1" calcOnSave="1"/>\n</workbook>');
    }

    zip.file(workbookPath, workbookXml);
  }

  const output = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(targetPath, output);
  return targetPath;
}

module.exports = {
  createDocument,
  buildDestinationPath,
  pickTemplatePath,
  deleteDocument,
  relocateBusinessDocuments,
  normalizeTemplateFile,
  __private: {
    applyAhmenTemplateWithZip,
    replaceAhmenNumericCell,
    replaceAhmenStringCell,
    buildNumericCell,
    formatNumericForCell
  }
};

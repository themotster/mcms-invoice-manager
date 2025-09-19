const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const db = require('./db');

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
  const eventName = sanitizeForFilename(event?.event_name || `${docType || 'document'}`);

  const numberPart = number ? `#${number} – ` : '';
  const fileNameParts = [iso, numberPart + clientName, eventName, human].filter(Boolean);
  const fileName = `${fileNameParts.join(' – ')}${extension}`;
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
    cell.value = value ?? '';
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

function fillAhmenTemplate(workbook, context) {
  const sheet = workbook.getWorksheet('Client Data') || workbook.worksheets[0];
  if (!sheet) return;

  // Client details
  setCellIfExists(sheet, 'B3', context.client?.name || '');
  setCellIfExists(sheet, 'B4', context.client?.email || '');
  setCellIfExists(sheet, 'B5', context.client?.phone || '');
  setCellIfExists(sheet, 'B6', context.client?.address1 || '');
  setCellIfExists(sheet, 'B7', context.client?.address2 || '');
  setCellIfExists(sheet, 'B8', context.client?.address3 || '');
  setCellIfExists(sheet, 'B9', context.client?.town || '');
  setCellIfExists(sheet, 'B10', context.client?.postcode || '');

  // Event details
  setCellIfExists(sheet, 'B13', context.event?.type || '');
  setCellIfExists(sheet, 'B14', context.event?.event_date ? formatDateParts(context.event.event_date).human : '');
  setCellIfExists(sheet, 'B15', context.event?.startTime || '');
  setCellIfExists(sheet, 'B16', context.event?.endTime || '');

  // Venue
  setCellIfExists(sheet, 'B19', context.event?.venue_name || '');
  setCellIfExists(sheet, 'B20', context.event?.venue_address1 || '');
  setCellIfExists(sheet, 'B21', context.event?.venue_address2 || '');
  setCellIfExists(sheet, 'B22', context.event?.venue_address3 || '');
  setCellIfExists(sheet, 'B23', context.event?.venue_town || '');
  setCellIfExists(sheet, 'B24', context.event?.venue_postcode || '');

  // Billing
  setCellIfExists(sheet, 'B27', context.totalAmount || '');
  setCellIfExists(sheet, 'B28', context.extraFees || '');
  setCellIfExists(sheet, 'B29', context.productionFees || '');
  setCellIfExists(sheet, 'B30', context.deposit || '');
  setCellIfExists(sheet, 'B31', context.balanceAmount || '');
  setCellIfExists(sheet, 'B32', context.balanceDate || '');
  setCellIfExists(sheet, 'B33', context.balanceRemind || '');

  // Other
  setCellIfExists(sheet, 'B36', context.serviceType || '');
  setCellIfExists(sheet, 'B37', context.specialistSingers || '');
}

async function generateExcelDocument({ templatePath, destinationPath, business, docType, context }) {
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

  if (business?.business_name === 'Motti Cohen Music Services' && docType === 'invoice') {
    const sheet = workbook.worksheets[0];
    if (sheet) {
      fillMcmsInvoiceWorksheet(sheet, context);
    }
  } else if (business?.business_name === 'AhMen A Cappella Ltd') {
    fillAhmenTemplate(workbook, context);
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

async function createDocument(documentData) {
  if (!documentData?.doc_type) {
    throw new Error('Document type is required');
  }
  if (!documentData?.business_id) {
    throw new Error('business_id is required to create a document');
  }

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
  const insertPayload = { ...documentData };
  delete insertPayload.file_path;
  const sessionIds = Array.isArray(documentData.session_ids) ? documentData.session_ids.slice() : [];
  delete insertPayload.session_ids;
  const lineItems = Array.isArray(documentData.line_items) ? documentData.line_items.slice() : [];
  delete insertPayload.line_items;

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

  const templatePath = pickTemplatePath(business, documentData.doc_type);
  const destinationPath = manualFilePath || buildDestinationPath({
    business,
    client,
    event,
    docType: documentData.doc_type,
    number: insertResult.number,
    templatePath
  });

  if (manualFilePath) {
    ensureDirectoryExists(path.dirname(manualFilePath));
  }

  const context = {
    business,
    client,
    event,
    docType: documentData.doc_type,
    number: insertResult.number,
    documentDate: documentData.document_date || new Date().toISOString(),
    dueDate: insertPayload.due_date,
    totalAmount: insertPayload.total_amount,
    balanceDue: insertPayload.balance_due,
    lineItems
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
    file_path: destinationPath
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

module.exports = {
  createDocument,
  buildDestinationPath,
  pickTemplatePath,
  deleteDocument
};

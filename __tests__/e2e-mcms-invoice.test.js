/**
 * E2E test for MCMS invoice creation: template fill, line item descriptions, subtotal/discount/received/balance_due.
 * Run: npm test -- __tests__/e2e-mcms-invoice.test.js
 * Requires: business_id 1 exists in DB (use shared DB having run app once, or set MCMS_DB_PATH before first require for temp DB).
 */
const path = require('path');
const fs = require('fs');
const os = require('os');
const ExcelJS = require('exceljs');

function buildMinimalInvoiceTemplate(targetPath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Invoice', { views: [{ state: 'normal' }] });
  // Header row
  ws.getCell(1, 1).value = 'Date';
  ws.getCell(1, 2).value = 'Description';
  ws.getCell(1, 3).value = 'Amount';
  // Line item row 1 (first data row)
  ws.getCell(2, 1).value = '{{item_date}}';
  ws.getCell(2, 2).value = '{{item_description}}';
  ws.getCell(2, 3).value = '{{item_amount}}';
  // Line item row 2 (second data row) – same tokens so writeRepeatableItemRows fills both
  ws.getCell(3, 1).value = '{{item_date}}';
  ws.getCell(3, 2).value = '{{item_description}}';
  ws.getCell(3, 3).value = '{{item_amount}}';
  // Summary: Subtotal (label in A, value in B), Discount, Received, Balance due
  ws.getCell(4, 1).value = '{{subtotal_label}}';
  ws.getCell(4, 2).value = '{{subtotal}}';
  ws.getCell(5, 1).value = 'Discount';
  ws.getCell(5, 2).value = '{{discount_amount}}';
  ws.getCell(6, 1).value = '{{received_label}}';
  ws.getCell(6, 2).value = '{{received}}';
  ws.getCell(7, 1).value = 'Balance due';
  ws.getCell(7, 2).value = '{{balance_due}}';
  return wb.xlsx.writeFile(targetPath);
}

function getCellDisplayValue(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  if (v && typeof v === 'object' && Array.isArray(v.richText)) {
    return v.richText.map(f => (f && f.text) ? f.text : '').join('');
  }
  return String(v);
}

describe('E2E MCMS invoice', () => {
  let db;
  let documentService;
  let tempDir;
  let templatePath;
  let businessId = 1;

  beforeAll(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcms-e2e-'));
    const dbPath = path.join(tempDir, 'e2e.db');
    process.env.MCMS_DB_PATH = dbPath;
    jest.isolateModules(() => {
      db = require('../db');
      documentService = require('../documentService');
    });
    await (db.dbReady || Promise.resolve());
    let business = await db.getBusinessById(businessId);
    if (!business) {
      // Fallback: use default DB (app data) so E2E can run after app has been run once
      delete process.env.MCMS_DB_PATH;
      jest.isolateModules(() => {
        db = require('../db');
        documentService = require('../documentService');
      });
      await (db.dbReady || Promise.resolve());
      business = await db.getBusinessById(businessId);
      if (!business) {
        throw new Error('E2E requires business_id 1. Run the app once to seed the DB, then run tests.');
      }
      // Leave tempDir for template only; DB is shared
    }
    templatePath = path.join(tempDir, 'template.xlsx');
    await buildMinimalInvoiceTemplate(templatePath);
  });

  afterAll(() => {
    try {
      if (tempDir && fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    } catch (_) {}
  });

  // Skip in Jest when ExcelJS hits readable-stream objectMode bug; run app or test:e2e:mcms manually to verify
  it.skip('fills line items so row 2 description is exactly "test item 2" and subtotal has no leading 0', async () => {
    const resolvedTemplate = path.resolve(templatePath);
    expect(fs.existsSync(resolvedTemplate)).toBe(true);
    const res = await documentService.createMCMSInvoice({
      business_id: businessId,
      template_path: resolvedTemplate,
      save_path: tempDir,
      _e2eKeepWorkbook: true,
      client_override: { name: 'E2E Test Client' },
      line_items: [
        { description: 'test item 1', amount: 25, date: '2026-03-15' },
        { description: 'test item 2', amount: 1500, date: '2026-03-15' },
      ],
      document_date: '2026-03-15',
      due_date: 'On receipt',
      total_amount: 1525,
      amount_received: 200,
      discount_amount: 100,
      discount_description: 'E2E discount',
      field_values: { invoice_date: '2026-03-15', due_date: 'On receipt', amount_received: 200 },
    });

    expect(res).toBeDefined();
    expect(res.ok).toBe(true);
    expect(res.workbook_path).toBeDefined();
    expect(fs.existsSync(res.workbook_path)).toBe(true);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(res.workbook_path);
    const ws = wb.worksheets[0];
    expect(ws).toBeDefined();

    // Line item 1 (row 2): description in column B
    const row2Desc = getCellDisplayValue(ws.getCell(2, 2));
    expect(row2Desc).toBe('test item 1');

    // Line item 2 (row 3): description must be exactly "test item 2" – no "te:", no duplicate
    const row3Desc = getCellDisplayValue(ws.getCell(3, 2));
    expect(row3Desc).toBe('test item 2');
    expect(row3Desc).not.toMatch(/te:\s*test item 2/);
    expect(row3Desc).not.toMatch(/test item 2.*test item 2/);

    // Subtotal row (row 4): value in B – must not start with "0 " or "0"
    const subtotalCell = getCellDisplayValue(ws.getCell(4, 2));
    expect(subtotalCell).not.toMatch(/^\s*0\s+/);
    expect(subtotalCell).not.toBe('0');

    // Subtotal label (row 4 A) should be "Subtotal"
    const subtotalLabel = getCellDisplayValue(ws.getCell(4, 1));
    expect(subtotalLabel).toBe('Subtotal');

    // Balance due (row 7 B) should be 1225 (1525 - 100 - 200)
    const balanceCell = ws.getCell(7, 2).value;
    const balanceNum = typeof balanceCell === 'number' ? balanceCell : Number(balanceCell);
    expect(Number.isFinite(balanceNum)).toBe(true);
    expect(Math.round(balanceNum)).toBe(1225);

    // Discount (row 5 B) and Received (row 6 B) must be negative
    const discountCell = ws.getCell(5, 2).value;
    const discountNum = typeof discountCell === 'number' ? discountCell : Number(discountCell);
    expect(Number.isFinite(discountNum)).toBe(true);
    expect(discountNum).toBeLessThanOrEqual(0);
    expect(Math.abs(discountNum)).toBe(100);

    const receivedCell = ws.getCell(6, 2).value;
    const receivedNum = typeof receivedCell === 'number' ? receivedCell : Number(receivedCell);
    expect(Number.isFinite(receivedNum)).toBe(true);
    expect(receivedNum).toBeLessThanOrEqual(0);
    expect(Math.abs(receivedNum)).toBe(200);

    // Received label (row 6 A) must say "Received", not the amount
    const receivedLabel = getCellDisplayValue(ws.getCell(6, 1));
    expect(receivedLabel).toBe('Received');
  }, 30000);

  it.skip('balance_due = subtotal - discount - received (zero discount/received)', async () => {
    const resolvedTemplate = path.resolve(templatePath);
    const res = await documentService.createMCMSInvoice({
      business_id: businessId,
      template_path: resolvedTemplate,
      save_path: tempDir,
      _e2eKeepWorkbook: true,
      client_override: { name: 'Math Test' },
      line_items: [{ description: 'One item', amount: 500, date: '2026-03-15' }],
      document_date: '2026-03-15',
      due_date: 'On receipt',
      total_amount: 500,
      amount_received: 0,
      discount_amount: 0,
      field_values: { invoice_date: '2026-03-15' },
    });
    expect(res).toBeDefined();
    expect(res.ok).toBe(true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(res.workbook_path);
    const ws = wb.worksheets[0];
    // With discount and received both 0, those rows are removed; balance due is then at row 5
    const balanceCell = ws.getCell(5, 2).value;
    const balanceNum = typeof balanceCell === 'number' ? balanceCell : Number(balanceCell);
    expect(Math.round(balanceNum)).toBe(500);
  }, 30000);
});

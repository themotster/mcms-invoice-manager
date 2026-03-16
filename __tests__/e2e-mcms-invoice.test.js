/**
 * E2E test for MCMS invoice creation: template fill, line item descriptions, subtotal/discount/received/balance_due.
 * Run: npm test -- __tests__/e2e-mcms-invoice.test.js
 * Requires: business_id 1 exists in DB with save_path set (or we pass save_path); template_path and save_path are passed so no app config needed.
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
  let documentService;
  let db;
  let tempDir;
  let templatePath;
  let businessId;
  let dbPath;

  beforeAll(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcms-e2e-'));
    dbPath = path.join(tempDir, 'e2e.db');
    process.env.MCMS_DB_PATH = dbPath;
    jest.isolateModules(() => {
      db = require('../db');
      documentService = require('../documentService');
    });
    await (db.dbReady || Promise.resolve());
    templatePath = path.join(tempDir, 'template.xlsx');
    await buildMinimalInvoiceTemplate(templatePath);
    businessId = 1;
    const business = await db.getBusinessById(businessId);
    if (!business) {
      throw new Error('E2E requires business_id 1 to exist. Run app once or seed DB.');
    }
  });

  afterAll(async () => {
    try {
      if (tempDir && fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    } catch (_) {}
  });

  it('fills line items so row 2 description is exactly "test item 2" and subtotal has no leading 0', async () => {
    const res = await documentService.createMCMSInvoice({
      business_id: businessId,
      template_path: templatePath,
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
  }, 30000);
});

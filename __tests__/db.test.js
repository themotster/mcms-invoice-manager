/**
 * DB layer tests: path, dbReady, getDocuments, getBusinessById, businessSettings.
 * Uses shared app DB (same as dev/packaged app) so business_id 1 exists after app or db-path test has run.
 */
const path = require('path');
const os = require('os');

describe('DB', () => {
  let db;

  beforeAll(async () => {
    delete process.env.MCMS_DB_PATH;
    delete process.env.INVOICE_MASTER_DB_PATH;
    jest.resetModules();
    db = require('../db');
    await (db.dbReady || Promise.resolve());
  });

  it('getDbPath returns shared app path when no env override', () => {
    const expected = path.join(os.homedir(), 'Library', 'Application Support', 'MCMS Invoice Manager', 'invoice_master.db');
    expect(db.getDbPath()).toBe(expected);
  });

  it('dbReady resolves', async () => {
    await expect(db.dbReady).resolves.toBeUndefined();
  });

  it('getBusinessById(1) returns business with save_path', async () => {
    const business = await db.getBusinessById(1);
    if (!business) {
      throw new Error('DB tests require business_id 1. Run the app once to seed the shared DB.');
    }
    expect(business).toHaveProperty('id', 1);
    expect(business).toHaveProperty('business_name');
    expect(business).toHaveProperty('save_path');
    expect(business).toHaveProperty('last_invoice_number');
  });

  it('businessSettings() returns array of business rows', async () => {
    const settings = await db.businessSettings();
    expect(Array.isArray(settings)).toBe(true);
    if (settings.length === 0) {
      throw new Error('DB tests require at least one business. Run the app once to seed the shared DB.');
    }
    const first = settings[0];
    expect(first).toHaveProperty('id');
    expect(first).toHaveProperty('business_name');
    expect(first).toHaveProperty('save_path');
  });

  it('getDocuments({ businessId: 1, docType: "invoice" }) returns array', async () => {
    const docs = await db.getDocuments({ businessId: 1, docType: 'invoice' });
    expect(Array.isArray(docs)).toBe(true);
  });

  it('getDocuments({ businessId: 1, docType: "Invoice" }) returns array (case-insensitive)', async () => {
    const docs = await db.getDocuments({ businessId: 1, docType: 'Invoice' });
    expect(Array.isArray(docs)).toBe(true);
  });

  it('getMergeFields() returns array', async () => {
    const fields = await db.getMergeFields();
    expect(Array.isArray(fields)).toBe(true);
    const hasClientName = fields.some(f => (f.field_key || '').toLowerCase() === 'client_name');
    expect(hasClientName).toBe(true);
  });

  it('getMaxInvoiceNumber(1) returns number or null', async () => {
    const max = await db.getMaxInvoiceNumber(1);
    expect(max === null || Number.isInteger(max)).toBe(true);
  });
});

/**
 * Regression test: dev and packaged app must use the same DB path by default.
 * Prevents data loss when switching between dev and released app.
 */
const path = require('path');
const os = require('os');

describe('DB path', () => {
  beforeEach(() => {
    delete process.env.MCMS_DB_PATH;
    delete process.env.INVOICE_MASTER_DB_PATH;
    jest.resetModules();
  });

  it('uses shared path when no env override (dev mode)', () => {
    const expected = path.join(
      os.homedir(),
      'Library',
      'Application Support',
      'MCMS Invoice Manager',
      'invoice_master.db'
    );
    const db = require('../db');
    expect(db.getDbPath()).toBe(expected);
  });

  it('uses env path when MCMS_DB_PATH is set (e.g. tests)', () => {
    const testPath = '/tmp/mcms-test.db';
    process.env.MCMS_DB_PATH = testPath;
    jest.resetModules();
    const db = require('../db');
    expect(db.getDbPath()).toBe(testPath);
  });
});

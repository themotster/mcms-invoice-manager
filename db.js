const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const os = require('os');

const settings = JSON.parse(fs.readFileSync(path.join(__dirname, 'settings.json')));
const sharedSupportDir = path.join(os.homedir(), 'Library', 'Application Support', 'AhMen Booking Manager');
const sharedDbPath = path.join(sharedSupportDir, 'invoice_master.db');

const isPackaged = (() => {
  try {
    // Only available in Electron main process.
    // eslint-disable-next-line global-require
    const { app } = require('electron');
    return Boolean(app && app.isPackaged);
  } catch (_err) {
    return false;
  }
})();

const envDbPath = process.env.AHMEN_DB_PATH || process.env.INVOICE_MASTER_DB_PATH;
const SQLITE_HEADER = Buffer.from('SQLite format 3\u0000');

const normalizeDbPath = (value) => {
  if (!value) return null;
  const trimmed = String(value).trim();
  if (!trimmed) return null;
  if (trimmed === ':memory:') return trimmed;
  return path.resolve(trimmed);
};

const isValidSqliteFile = (filePath) => {
  try {
    const stats = fs.statSync(filePath);
    if (!stats.isFile()) return false;
    if (stats.size < SQLITE_HEADER.length) return false;
    const fd = fs.openSync(filePath, 'r');
    const buffer = Buffer.alloc(SQLITE_HEADER.length);
    fs.readSync(fd, buffer, 0, SQLITE_HEADER.length, 0);
    fs.closeSync(fd);
    return buffer.equals(SQLITE_HEADER);
  } catch (_err) {
    return false;
  }
};

if (isPackaged) {
  try {
    fs.mkdirSync(sharedSupportDir, { recursive: true });
    const legacyPath = normalizeDbPath(settings.db_path);
    if (legacyPath && fs.existsSync(legacyPath) && isValidSqliteFile(legacyPath) && !fs.existsSync(sharedDbPath)) {
      fs.copyFileSync(legacyPath, sharedDbPath);
      console.log(`Migrated database to ${sharedDbPath}`);
    }
  } catch (err) {
    console.error('Failed to migrate database', err);
  }
}

const candidatePaths = [];
const normalizedEnv = normalizeDbPath(envDbPath);
const normalizedSettings = normalizeDbPath(settings.db_path);
const pushCandidate = (value) => {
  if (!value) return;
  if (!candidatePaths.includes(value)) candidatePaths.push(value);
};

pushCandidate(normalizedEnv);
if (isPackaged) pushCandidate(sharedDbPath);
pushCandidate(normalizedSettings);
if (!candidatePaths.length) pushCandidate(sharedDbPath);

const ensureDbFile = (targetPath) => {
  if (!targetPath || targetPath === ':memory:') return true;
  try {
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    if (fs.existsSync(targetPath) && !isValidSqliteFile(targetPath)) {
      const backupPath = `${targetPath}.corrupt-${Date.now()}`;
      try {
        fs.renameSync(targetPath, backupPath);
        console.warn(`Moved invalid database to ${backupPath}`);
      } catch (err) {
        console.error('Failed to move invalid database', err);
        return false;
      }
    }
    return true;
  } catch (err) {
    console.error(`DB path not writable: ${targetPath}`, err);
    return false;
  }
};

let dbPath = candidatePaths.find(p => ensureDbFile(p)) || ':memory:';

const db = new sqlite3.Database(dbPath, sqlite3.OPEN_READWRITE | sqlite3.OPEN_CREATE, err => {
  if (err) {
    console.error('Failed to open database', err);
  }
});
db.configure('busyTimeout', 5000);
const mergeFieldDefaults = require('./config/mergeFields.json');

function escapeLikePattern(value) {
  if (typeof value !== 'string') return value;
  return value
    .replace(/\\/g, '\\\\')
    .replace(/%/g, '\\%')
    .replace(/_/g, '\\_');
}

function toSqliteDateTime(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.valueOf())) return null;
  return date.toISOString().slice(0, 19).replace('T', ' ');
}

const DEFAULT_BUSINESS_ID = 1;
const DEFAULT_BUSINESSES = [
  {
    id: 1,
    business_name: 'Motti Cohen Music Services',
    last_invoice_number: 704,
    last_quote_number: 0,
    save_path: '/Users/Shared/Invoices/MCMS'
  },
  {
    id: 2,
    business_name: 'AhMen A Cappella Ltd',
    last_invoice_number: 882,
    last_quote_number: 0,
    save_path: '/Users/Shared/Invoices/AhMen'
  }
];

const DEFAULT_DOCUMENT_DEFINITIONS = [
  {
    key: 'quote',
    doc_type: 'quote',
    label: 'Quote',
    description: 'Quote document with pricing totals.',
    invoice_variant: null,
    is_primary: 0,
    sort_order: 0,
    locked: 1
  },
  {
    key: 'contract',
    doc_type: 'contract',
    label: 'Contract',
    description: 'Contract ready for signatures.',
    invoice_variant: null,
    is_primary: 0,
    sort_order: 1,
    locked: 1
  },
  {
    key: 'invoice_deposit',
    doc_type: 'invoice',
    label: 'Invoice – Deposit',
    description: 'Deposit invoice for the booking.',
    invoice_variant: 'deposit',
    is_primary: 0,
    sort_order: 2,
    locked: 1
  },
  {
    key: 'invoice_balance',
    doc_type: 'invoice',
    label: 'Invoice – Balance',
    description: 'Balance invoice for the booking.',
    invoice_variant: 'balance',
    is_primary: 0,
    sort_order: 3,
    locked: 1
  }
];

const DEFAULT_FIELD_VALUE_SOURCES = {
  client_name: 'jobsheet.client_name',
  client_email: 'jobsheet.client_email',
  client_phone: 'jobsheet.client_phone',
  client_address: 'jobsheet.client_address',
  client_address1: 'jobsheet.client_address1',
  client_address2: 'jobsheet.client_address2',
  client_address3: 'jobsheet.client_address3',
  client_town: 'jobsheet.client_town',
  client_postcode: 'jobsheet.client_postcode',
  event_type: 'jobsheet.event_type',
  event_date: 'jobsheet.event_date',
  event_start: 'jobsheet.event_start',
  event_end: 'jobsheet.event_end',
  venue_address: 'jobsheet.venue_address',
  venue_name: 'jobsheet.venue_name',
  venue_address1: 'jobsheet.venue_address1',
  venue_address2: 'jobsheet.venue_address2',
  venue_address3: 'jobsheet.venue_address3',
  venue_town: 'jobsheet.venue_town',
  venue_postcode: 'jobsheet.venue_postcode',
  caterer_name: 'jobsheet.caterer_name',
  ahmen_fee: 'jobsheet.ahmen_fee',
  vat_amount: 'jobsheet.vat_amount',
  total_amount: 'context.totalAmount',
  production_fees: 'context.productionFees',
  deposit_amount: 'context.depositAmount',
  balance_amount: 'context.balanceAmount',
  balance_due_date: 'context.balanceDate',
  balance_reminder_date: 'context.balanceRemind',
  service_types: 'jobsheet.service_types',
  specialist_singers: 'jobsheet.specialist_singers',
  special_conditions: 'jobsheet.special_conditions'
};

const TRASH_DIR_NAME = '.trash';
const MAX_TREE_DEPTH = 6;
const MAX_TREE_ENTRIES = 4000;

const BUSINESS_SETTINGS_MUTABLE_FIELDS = new Set([
  'save_path',
  'last_invoice_number',
  'last_quote_number'
]);

const AHMEN_JOBSHEET_FIELDS = [
  'business_id',
  'status',
  'client_name',
  'client_email',
  'client_phone',
  'client_address',
  'client_address1',
  'client_address2',
  'client_address3',
  'client_town',
  'client_postcode',
  'event_type',
  'event_date',
  'event_start',
  'event_end',
  'venue_id',
  'venue_name',
  'venue_address',
  'venue_address1',
  'venue_address2',
  'venue_address3',
  'venue_town',
  'venue_postcode',
  'venue_same_as_client',
  'ahmen_fee',
  'production_fees',
  'vat_enabled',
  'vat_amount',
  'deposit_amount',
  'balance_amount',
  'balance_due_date',
  'balance_reminder_date',
  'service_types',
  'specialist_singers',
  'special_conditions',
  'gig_info',
  'notes',
  'pricing_service_id',
  'pricing_selected_singers',
  'pricing_discount',
  'pricing_discount_type',
  'pricing_discount_value',
  'pricing_production_items',
  'pricing_production_subtotal',
  'pricing_production_discount',
  'pricing_production_discount_type',
  'pricing_production_discount_value',
  'pricing_production_total',
  'pricing_total'
];

const AHMEN_JOBSHEET_NUMERIC_FIELDS = new Set([
  'ahmen_fee',
  'production_fees',
  'vat_amount',
  'deposit_amount',
  'balance_amount',
  'pricing_discount',
  'pricing_discount_value',
  'pricing_production_subtotal',
  'pricing_production_discount_value',
  'pricing_production_total',
  'pricing_total'
]);

const AHMEN_JOBSHEET_BOOLEAN_FIELDS = new Set([
  'venue_same_as_client',
  'vat_enabled'
]);

const AHMEN_JOBSHEET_INTEGER_FIELDS = new Set([
  'venue_id'
]);

const AHMEN_JOBSHEET_STATUS_VALUES = new Set(['enquiry', 'quoted', 'contracting', 'confirmed', 'completed']);

const AHMEN_VENUE_FIELDS = [
  'business_id',
  'name',
  'address1',
  'address2',
  'address3',
  'town',
  'postcode',
  'is_private'
];

const MERGE_FIELD_TABLE = 'merge_fields';
const MERGE_FIELD_BINDINGS_TABLE = 'merge_field_bindings';
const MERGE_FIELD_BINDING_UNIQUE_COLUMNS = ['field_key', 'template', 'sheet', 'cell'];

function logDuplicateColumn(err) {
  if (!err) return;
  const duplicateMsg = 'duplicate column name';
  if (err.message && err.message.toLowerCase().includes(duplicateMsg)) return;
  console.error('SQLite schema migration error:', err.message || err);
}

function initializeDatabase() {
  db.serialize(() => {
    db.run('PRAGMA foreign_keys = ON');

    db.run(`CREATE TABLE IF NOT EXISTS business_settings (
      id INTEGER PRIMARY KEY,
      business_name TEXT NOT NULL UNIQUE,
      last_invoice_number INTEGER DEFAULT 0,
      last_quote_number INTEGER DEFAULT 0,
      save_path TEXT NOT NULL
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS ${MERGE_FIELD_TABLE} (
      field_key TEXT PRIMARY KEY,
      label TEXT NOT NULL,
      placeholder TEXT,
      category TEXT,
      description TEXT,
      show_in_jobsheet INTEGER DEFAULT 1,
      active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS ${MERGE_FIELD_BINDINGS_TABLE} (
      binding_id INTEGER PRIMARY KEY AUTOINCREMENT,
      field_key TEXT NOT NULL,
      template TEXT NOT NULL,
      sheet TEXT,
      cell TEXT,
      data_type TEXT DEFAULT 'string',
      style TEXT,
      format TEXT,
      FOREIGN KEY(field_key) REFERENCES ${MERGE_FIELD_TABLE}(field_key)
    )`);

    db.run(`CREATE UNIQUE INDEX IF NOT EXISTS idx_merge_field_bindings_unique
      ON ${MERGE_FIELD_BINDINGS_TABLE} (${MERGE_FIELD_BINDING_UNIQUE_COLUMNS.join(', ')})`);

    // Backfill timestamps for merge field bindings if columns don't exist
    db.run(`ALTER TABLE ${MERGE_FIELD_BINDINGS_TABLE} ADD COLUMN created_at TEXT`, logDuplicateColumn);
    db.run(`ALTER TABLE ${MERGE_FIELD_BINDINGS_TABLE} ADD COLUMN updated_at TEXT`, logDuplicateColumn);

    db.run(`CREATE TABLE IF NOT EXISTS merge_field_value_sources (
      field_key TEXT PRIMARY KEY,
      source_type TEXT NOT NULL,
      source_path TEXT,
      literal_value TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY(field_key) REFERENCES ${MERGE_FIELD_TABLE}(field_key) ON DELETE CASCADE
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS businesses (
      business_id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      branding_template TEXT,
      settings TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS clients (
      client_id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER,
      name TEXT NOT NULL,
      email TEXT,
      phone TEXT,
      address TEXT,
      address1 TEXT,
      address2 TEXT,
      town TEXT,
      postcode TEXT,
      contact TEXT,
      FOREIGN KEY (business_id) REFERENCES business_settings(id)
    )`);

    // Normalized client contact details for multiple values per client
    db.run(`CREATE TABLE IF NOT EXISTS client_emails (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      client_id INTEGER NOT NULL,
      label TEXT,
      email TEXT NOT NULL,
      is_primary INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (client_id) REFERENCES clients(client_id) ON DELETE CASCADE
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS client_phones (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      client_id INTEGER NOT NULL,
      label TEXT,
      phone TEXT NOT NULL,
      is_primary INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (client_id) REFERENCES clients(client_id) ON DELETE CASCADE
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS client_addresses (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      client_id INTEGER NOT NULL,
      label TEXT,
      address1 TEXT,
      address2 TEXT,
      town TEXT,
      postcode TEXT,
      country TEXT,
      is_primary INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (client_id) REFERENCES clients(client_id) ON DELETE CASCADE
    )`);
    db.run('CREATE INDEX IF NOT EXISTS idx_client_emails_client ON client_emails (client_id)');
    db.run('CREATE INDEX IF NOT EXISTS idx_client_phones_client ON client_phones (client_id)');
    db.run('CREATE INDEX IF NOT EXISTS idx_client_addresses_client ON client_addresses (client_id)');

    db.run(`CREATE TABLE IF NOT EXISTS events (
      event_id INTEGER PRIMARY KEY AUTOINCREMENT,
      client_id INTEGER NOT NULL,
      business_id INTEGER,
      event_name TEXT NOT NULL,
      event_date TEXT,
      venue_name TEXT,
      venue_address1 TEXT,
      venue_address2 TEXT,
      town TEXT,
      postcode TEXT,
      notes TEXT,
      FOREIGN KEY (client_id) REFERENCES clients(client_id),
      FOREIGN KEY (business_id) REFERENCES business_settings(id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS documents (
      document_id INTEGER PRIMARY KEY AUTOINCREMENT,
      event_id INTEGER,
      jobsheet_id INTEGER,
      business_id INTEGER,
      doc_type TEXT NOT NULL,
      number INTEGER,
      status TEXT,
      total_amount REAL,
      balance_due REAL,
      due_date TEXT,
      file_path TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (event_id) REFERENCES events(event_id),
      FOREIGN KEY (jobsheet_id) REFERENCES ahmen_jobsheets(jobsheet_id),
      FOREIGN KEY (business_id) REFERENCES business_settings(id)
    )`);

    // Line items for invoices/quotes
    db.run(`CREATE TABLE IF NOT EXISTS document_items (
      item_id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      item_type TEXT,
      description TEXT,
      quantity REAL,
      unit TEXT,
      rate REAL,
      amount REAL,
      sort_order INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (document_id) REFERENCES documents(document_id) ON DELETE CASCADE
    )`);
    db.run('CREATE INDEX IF NOT EXISTS idx_document_items_doc ON document_items (document_id, sort_order)');

    // Email log for in-app sending
    db.run(`CREATE TABLE IF NOT EXISTS email_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER,
      jobsheet_id INTEGER,
      to_address TEXT NOT NULL,
      cc_address TEXT,
      bcc_address TEXT,
      subject TEXT,
      body TEXT,
      attachments TEXT,
      provider TEXT DEFAULT 'graph',
      status TEXT DEFAULT 'sent',
      message_id TEXT,
      sent_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      FOREIGN KEY (jobsheet_id) REFERENCES ahmen_jobsheets(jobsheet_id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS scheduled_emails (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      email_log_id INTEGER,
      business_id INTEGER,
      jobsheet_id INTEGER,
      to_address TEXT NOT NULL,
      cc_address TEXT,
      bcc_address TEXT,
      subject TEXT,
      body TEXT,
      attachments TEXT,
      is_html INTEGER DEFAULT 1,
      send_at TEXT NOT NULL,
      status TEXT DEFAULT 'pending',
      attempt_count INTEGER DEFAULT 0,
      last_error TEXT,
      sent_at TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (email_log_id) REFERENCES email_log(id) ON DELETE SET NULL,
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      FOREIGN KEY (jobsheet_id) REFERENCES ahmen_jobsheets(jobsheet_id)
    )`);

    db.run(`CREATE INDEX IF NOT EXISTS idx_scheduled_emails_status_sendat
      ON scheduled_emails (status, send_at)`);

    db.run('ALTER TABLE scheduled_emails ADD COLUMN is_html INTEGER DEFAULT 1', logDuplicateColumn);

    db.run(`CREATE TABLE IF NOT EXISTS planner_actions (
      action_id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER NOT NULL,
      jobsheet_id INTEGER NOT NULL,
      action_key TEXT NOT NULL,
      scheduled_for TEXT NOT NULL,
      status TEXT DEFAULT 'pending',
      completed_at TEXT,
      last_notified_at TEXT,
      last_email_at TEXT,
      last_error TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      FOREIGN KEY (jobsheet_id) REFERENCES ahmen_jobsheets(jobsheet_id),
      UNIQUE (business_id, jobsheet_id, action_key, scheduled_for)
    )`);

    db.run(`CREATE INDEX IF NOT EXISTS idx_planner_actions_status_date
      ON planner_actions (status, scheduled_for)`);

    db.run(`CREATE INDEX IF NOT EXISTS idx_planner_actions_business
      ON planner_actions (business_id, jobsheet_id)`);

  db.run(`CREATE TABLE IF NOT EXISTS document_definitions (
    definition_id INTEGER PRIMARY KEY AUTOINCREMENT,
    business_id INTEGER NOT NULL,
    key TEXT NOT NULL,
      doc_type TEXT NOT NULL,
      label TEXT NOT NULL,
      description TEXT,
      invoice_variant TEXT,
      template_path TEXT,
      is_active INTEGER DEFAULT 1,
      is_locked INTEGER DEFAULT 0,
      sort_order INTEGER DEFAULT 0,
      sheet_exports TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      UNIQUE (business_id, key)
    )`);

  db.run(`CREATE INDEX IF NOT EXISTS idx_document_definitions_business_sort
    ON document_definitions (business_id, sort_order)`);

  db.run(`CREATE TABLE IF NOT EXISTS document_definition_tombstones (
    business_id INTEGER NOT NULL,
    key TEXT NOT NULL,
    PRIMARY KEY (business_id, key)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS jobsheet_template_overrides (
      override_id INTEGER PRIMARY KEY AUTOINCREMENT,
      jobsheet_id INTEGER NOT NULL,
      definition_key TEXT NOT NULL,
      template_path TEXT NOT NULL,
      updated_at TEXT DEFAULT (datetime('now')),
      UNIQUE (jobsheet_id, definition_key)
    )`);

    db.run(`CREATE INDEX IF NOT EXISTS idx_jobsheet_template_overrides_jobsheet
      ON jobsheet_template_overrides (jobsheet_id)`);

    db.run(`CREATE TABLE IF NOT EXISTS event_musicians (
      musician_id INTEGER PRIMARY KEY AUTOINCREMENT,
      event_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      role TEXT,
      fee REAL DEFAULT 0,
      paid_status TEXT DEFAULT 'unpaid',
      FOREIGN KEY (event_id) REFERENCES events(event_id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS timekeeper_sessions (
      session_id INTEGER PRIMARY KEY AUTOINCREMENT,
      client_id INTEGER,
      event_id INTEGER,
      description TEXT,
      session_date TEXT,
      duration_minutes INTEGER,
      rate REAL,
      amount REAL,
      exported INTEGER DEFAULT 0,
      FOREIGN KEY (client_id) REFERENCES clients(client_id),
      FOREIGN KEY (event_id) REFERENCES events(event_id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS ahmen_venues (
      venue_id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      address1 TEXT,
      address2 TEXT,
      address3 TEXT,
      town TEXT,
      postcode TEXT,
      is_private INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id)
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS ahmen_jobsheets (
      jobsheet_id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER NOT NULL,
      status TEXT DEFAULT 'enquiry',
      client_name TEXT NOT NULL,
      client_email TEXT,
      client_phone TEXT,
      client_address TEXT,
      client_address1 TEXT,
      client_address2 TEXT,
      client_address3 TEXT,
      client_town TEXT,
      client_postcode TEXT,
      event_type TEXT,
      event_date TEXT,
      event_start TEXT,
      event_end TEXT,
      venue_id INTEGER,
      venue_name TEXT,
      venue_address TEXT,
      venue_address1 TEXT,
      venue_address2 TEXT,
      venue_address3 TEXT,
      venue_town TEXT,
      venue_postcode TEXT,
      caterer_name TEXT,
      venue_same_as_client INTEGER DEFAULT 0,
      ahmen_fee REAL,
      specialist_fees REAL,
      production_fees REAL,
      vat_amount REAL,
      deposit_amount REAL,
      balance_amount REAL,
      balance_due_date TEXT,
      balance_reminder_date TEXT,
      service_types TEXT,
      specialist_singers TEXT,
      special_conditions TEXT,
      gig_info TEXT,
      notes TEXT,
      pricing_service_id TEXT,
      pricing_selected_singers TEXT,
      pricing_custom_fees TEXT,
      pricing_discount REAL,
      pricing_discount_type TEXT,
      pricing_discount_value REAL,
      vat_enabled INTEGER DEFAULT 0,
      pricing_production_items TEXT,
      pricing_production_subtotal REAL,
      pricing_production_discount TEXT,
      pricing_production_discount_type TEXT,
      pricing_production_discount_value REAL,
      pricing_production_total REAL,
      pricing_total REAL,
      archived_at TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      FOREIGN KEY (venue_id) REFERENCES ahmen_venues(venue_id)
    )`);

    // Extend existing clients table with new columns if they do not yet exist
    db.run('ALTER TABLE clients ADD COLUMN contact TEXT', logDuplicateColumn);
    db.run('ALTER TABLE clients ADD COLUMN address1 TEXT', logDuplicateColumn);
    db.run('ALTER TABLE clients ADD COLUMN address2 TEXT', logDuplicateColumn);
    db.run('ALTER TABLE clients ADD COLUMN town TEXT', logDuplicateColumn);
    db.run('ALTER TABLE clients ADD COLUMN postcode TEXT', logDuplicateColumn);

    db.run('ALTER TABLE events ADD COLUMN business_id INTEGER', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN client_name TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN event_name TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN event_date TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN document_date TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN jobsheet_id INTEGER', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN updated_at TEXT DEFAULT (datetime(\'now\'))', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN definition_key TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN invoice_variant TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN reminder_date TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN reminder_sent_at TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN paid_at TEXT', logDuplicateColumn);
    db.run('ALTER TABLE documents ADD COLUMN is_locked INTEGER DEFAULT 0', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN caterer_name TEXT', logDuplicateColumn);
    db.run('ALTER TABLE document_definitions ADD COLUMN sheet_exports TEXT', logDuplicateColumn);

    // Template path columns removed; definitions manage their own template_path now.
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_discount_type TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_discount_value REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_items TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_subtotal REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_discount TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_discount_type TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_discount_value REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_production_total REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN venue_id INTEGER', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN venue_same_as_client INTEGER DEFAULT 0', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_service_id TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_selected_singers TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_custom_fees TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN vat_amount REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN vat_enabled INTEGER DEFAULT 0', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_discount REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_total REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN gig_info TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN archived_at TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN client_address TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN venue_address TEXT', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN special_conditions TEXT', logDuplicateColumn);
    
    db.run(`UPDATE ahmen_jobsheets SET status='enquiry' WHERE status IS NULL OR status='' OR status='draft'`, err => {
      if (err) console.error('Failed to normalize jobsheet status:', err);
    });

    seedBusinesses();
    syncLegacyBusinessesTable();
    seedMergeFieldDefaults();
    db.run(
      `UPDATE ${MERGE_FIELD_TABLE}
       SET placeholder = 'TOTAL_FEES', updated_at = datetime('now')
       WHERE field_key = 'total_amount' AND (placeholder IS NULL OR placeholder = '' OR placeholder = 'TOTAL')`
    );
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET cell = 'B30'
       WHERE field_key = 'total_amount' AND template = 'ahmen_excel' AND sheet = 'Client Data'`
    );
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET cell = 'B31'
       WHERE field_key = 'deposit_amount' AND template = 'ahmen_excel' AND sheet = 'Client Data'`
    );
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET cell = 'B32'
       WHERE field_key = 'balance_amount' AND template = 'ahmen_excel' AND sheet = 'Client Data'`
    );
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET cell = 'B33'
       WHERE field_key = 'balance_due_date' AND template = 'ahmen_excel' AND sheet = 'Client Data'`
    );
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET cell = 'B34'
       WHERE field_key = 'balance_reminder_date' AND template = 'ahmen_excel' AND sheet = 'Client Data'`
    );
    db.run(
      `INSERT OR IGNORE INTO ${MERGE_FIELD_BINDINGS_TABLE} (field_key, template, sheet, cell, data_type, style, format)
       VALUES ('ahmen_fee', 'ahmen_excel', 'Client Data', 'B27', 'number', '12', NULL)`
    );
    // Ensure numeric currency formatting on key fee fields (don't override if already set)
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET data_type = 'number', format = 'currency', updated_at = datetime('now')
       WHERE template = 'ahmen_excel' AND (data_type IS NULL OR data_type = '' OR data_type = 'string')
         AND field_key IN ('total_amount','deposit_amount','balance_amount','production_fees','extra_fees')`
    );
    // Ensure date data type on date fields
    db.run(
      `UPDATE ${MERGE_FIELD_BINDINGS_TABLE}
       SET data_type = 'date', updated_at = datetime('now')
       WHERE template = 'ahmen_excel' AND (data_type IS NULL OR data_type = '' OR data_type = 'string')
         AND field_key IN ('event_date','balance_due_date','balance_reminder_date')`
    );
    // Cleanup: remove legacy AHMEN_FEE bindings on Quote/Invoice F21 now that templates use curly placeholders
    db.run(
      `DELETE FROM ${MERGE_FIELD_BINDINGS_TABLE}
       WHERE field_key = 'ahmen_fee' AND template = 'ahmen_excel' AND cell = 'F21'
         AND sheet IN ('Quote','Quotes','QUOTE','Invoice','INVOICE','Invoice – Deposit','Invoice - Deposit','Invoice – Balance','Invoice - Balance')`
    );
    db.run(
      `DELETE FROM ${MERGE_FIELD_BINDINGS_TABLE}
       WHERE field_key = 'ahmen_fee' AND template = 'ahmen_excel' AND cell = 'F49'
         AND sheet IN ('Booking Schedule','Booking schedule')`
    );
    seedMergeFieldValueSources();
    migrateBusinessSettingsDropTemplateColumns(() => {
      migrateDocumentDefinitionsTable(() => {
        // Ensure deposit invoice can be reseeded if a tombstone was left behind
        reinstateDepositInvoiceDefinition();
        seedDocumentDefinitions();
        // Backfill deposit template path from balance if missing
        backfillDepositTemplatePathFromBalance();
        // Log quick summary of template_path coverage per business
        try {
          db.all(
            `SELECT business_id, COUNT(*) AS total, SUM(CASE WHEN template_path IS NULL OR template_path = '' THEN 1 ELSE 0 END) AS missing
             FROM document_definitions GROUP BY business_id`,
            [],
            (_err, rows) => {
              if (Array.isArray(rows)) {
                rows.forEach(r => {
                  if (r && Number(r.missing) > 0) {
                    console.log(`[defs] business ${r.business_id}: ${r.missing}/${r.total} templates missing path`);
                  }
                });
              }
            }
          );
        } catch (_err) {}
      });
    });
  });
}

function seedBusinesses() {
  const allowedNames = DEFAULT_BUSINESSES.map(b => b.business_name);

  DEFAULT_BUSINESSES.forEach(business => {
    db.run(
      `INSERT OR IGNORE INTO business_settings (
        id, business_name, last_invoice_number, last_quote_number, save_path
      )
       VALUES (?, ?, ?, ?, ?)`,
      [
        business.id,
        business.business_name,
        business.last_invoice_number,
        business.last_quote_number,
        business.save_path
      ]
    );

    db.run(
      `UPDATE business_settings
       SET business_name = ?, last_invoice_number = ?, last_quote_number = ?,
           save_path = CASE WHEN save_path IS NULL OR save_path = '' THEN ? ELSE save_path END
       WHERE id = ?`,
      [
        business.business_name,
        business.last_invoice_number,
        business.last_quote_number,
        business.save_path,
        business.id
      ]
    );
  });

  if (allowedNames.length) {
    const placeholders = allowedNames.map(() => '?').join(', ');
    db.run(
      `DELETE FROM business_settings WHERE business_name NOT IN (${placeholders})`,
      allowedNames
    );
  }
}

function seedMergeFieldDefaults() {
  if (!Array.isArray(mergeFieldDefaults) || !mergeFieldDefaults.length) return;

  db.serialize(() => {
    const insertField = db.prepare(
      `INSERT OR IGNORE INTO ${MERGE_FIELD_TABLE} (
        field_key, label, placeholder, category, description, show_in_jobsheet, active
      ) VALUES (?, ?, ?, ?, ?, ?, 1)`
    );

    const insertBinding = db.prepare(
      `INSERT OR IGNORE INTO ${MERGE_FIELD_BINDINGS_TABLE} (
        field_key, template, sheet, cell, data_type, style, format
      ) VALUES (?, ?, ?, ?, ?, ?, ?)`
    );

    mergeFieldDefaults.forEach(field => {
      insertField.run(
        field.field_key,
        field.label,
        field.placeholder || null,
        field.category || null,
        field.description || null,
        field.show_in_jobsheet ? 1 : 0,
        err => {
          if (err) console.error('Failed to seed merge field', field.field_key, err);
        }
      );

      if (Array.isArray(field.bindings)) {
        field.bindings.forEach(binding => {
          insertBinding.run(
            field.field_key,
            binding.template,
            binding.sheet || null,
            binding.cell || null,
            binding.data_type || 'string',
            binding.style || null,
            binding.format || null,
            err => {
              if (err) console.error('Failed to seed merge field binding', field.field_key, err);
            }
          );
        });
      }
    });

    insertField.finalize();
    insertBinding.finalize();
  });
}

function seedMergeFieldValueSources() {
  Object.entries(DEFAULT_FIELD_VALUE_SOURCES).forEach(([fieldKey, sourcePath]) => {
    if (!fieldKey || !sourcePath) return;
    db.run(
      `INSERT OR IGNORE INTO merge_field_value_sources (field_key, source_type, source_path, literal_value, created_at, updated_at)
       VALUES (?, 'contextPath', ?, NULL, datetime('now'), datetime('now'))`,
      [fieldKey, sourcePath],
      err => {
        if (err) {
          console.error('Failed to seed placeholder data mapping', fieldKey, err);
        }
      }
    );
  });
}

// Clear any tombstone that may hide the default deposit invoice definition
function reinstateDepositInvoiceDefinition() {
  try {
    db.all('SELECT id FROM business_settings', [], (err, businesses) => {
      if (err) return;
      const list = Array.isArray(businesses) ? businesses : [];
      list.forEach(biz => {
        const businessId = biz?.id;
        if (!Number.isInteger(businessId)) return;
        db.run(
          'DELETE FROM document_definition_tombstones WHERE business_id = ? AND key = ?;',
          [businessId, 'invoice_deposit'],
          () => {}
        );
      });
    });
  } catch (_err) { /* no-op */ }
}

// If the deposit definition has no template path but balance does, copy it across
function backfillDepositTemplatePathFromBalance() {
  try {
    db.all('SELECT id FROM business_settings', [], (err, businesses) => {
      if (err) return;
      const list = Array.isArray(businesses) ? businesses : [];
      list.forEach(biz => {
        const businessId = biz?.id;
        if (!Number.isInteger(businessId)) return;
        const sql = `UPDATE document_definitions AS dep
          SET template_path = (
            SELECT bal.template_path FROM document_definitions AS bal
            WHERE bal.business_id = dep.business_id AND bal.key = 'invoice_balance'
              AND bal.template_path IS NOT NULL AND bal.template_path <> ''
          ),
          updated_at = datetime('now')
          WHERE dep.business_id = ? AND dep.key = 'invoice_deposit'
            AND (dep.template_path IS NULL OR dep.template_path = '')
            AND EXISTS (
              SELECT 1 FROM document_definitions AS bal
              WHERE bal.business_id = dep.business_id AND bal.key = 'invoice_balance'
                AND bal.template_path IS NOT NULL AND bal.template_path <> ''
            )`;
        db.run(sql, [businessId], () => {});
      });
    });
  } catch (_err) { /* no-op */ }
}

// Business-level template defaults removed: definitions own template_path

function seedDocumentDefinitions() {
  db.all(`SELECT id FROM business_settings`, (err, businesses) => {
    if (err) {
      console.error('Failed to load businesses for document definition seeding', err);
      return;
    }

    businesses.forEach(business => {
      if (!business || business.id == null) return;
      const businessId = business.id;

      db.all(
        `SELECT key FROM document_definitions WHERE business_id = ?`,
        [businessId],
        (defErr, rows) => {
          if (defErr) {
            console.error('Failed to read document definitions', defErr);
            return;
          }

          const existingKeys = new Set(Array.isArray(rows) ? rows.map(row => row.key) : []);

          db.all(
            'SELECT key FROM document_definition_tombstones WHERE business_id = ?',
            [businessId],
            (tombErr, tombRows) => {
              if (tombErr) {
                console.error('Failed to read document definition tombstones', tombErr);
                return;
              }

              const tombstonedKeys = new Set(Array.isArray(tombRows) ? tombRows.map(row => row.key) : []);

              DEFAULT_DOCUMENT_DEFINITIONS.forEach(definition => {
                if (tombstonedKeys.has(definition.key)) {
                  return;
                }
              const templatePath = null;

            db.run(
              `INSERT OR IGNORE INTO document_definitions (
                business_id,
                key,
                doc_type,
                label,
                description,
                invoice_variant,
                template_path,
                is_active,
                is_locked,
                sort_order,
                created_at,
                updated_at
              ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?, datetime('now'), datetime('now'))`,
              [
                businessId,
                definition.key,
                definition.doc_type,
                definition.label,
                definition.description || null,
                definition.invoice_variant || null,
                templatePath || null,
                definition.locked ? 1 : 0,
                definition.sort_order
              ],
              insertErr => {
                if (insertErr) {
                  console.error('Failed to seed document definition', definition.key, insertErr);
                }
              }
            );

            const updateSql = `UPDATE document_definitions
              SET doc_type = ?,
                  label = ?,
                  description = ?,
                  invoice_variant = ?,
                  is_locked = CASE WHEN is_locked = 1 THEN 1 ELSE ? END,
                  template_path = CASE WHEN ? IS NOT NULL AND (template_path IS NULL OR template_path = '') THEN ? ELSE template_path END
              WHERE business_id = ? AND key = ?`;

            db.run(
              updateSql,
              [
                definition.doc_type,
                definition.label,
                definition.description || null,
                definition.invoice_variant || null,
                definition.locked ? 1 : 0,
                templatePath || null,
                templatePath || null,
                businessId,
                definition.key
              ],
              updateErr => {
                if (updateErr) {
                  console.error('Failed to update seeded document definition', definition.key, updateErr);
                }
              }
            );

            if (!existingKeys.has(definition.key)) {
              existingKeys.add(definition.key);
            }
              });
            }
          );
        }
      );
    });
  });
}

function migrateDocumentDefinitionsTable(done) {
  try {
    db.all('PRAGMA table_info(document_definitions)', (err, rows) => {
      if (err) {
        console.error('Failed to inspect document_definitions table', err);
        if (typeof done === 'function') done();
        return;
      }
      const columns = Array.isArray(rows) ? rows.map(r => r.name) : [];
      const hasFileSuffix = columns.includes('file_suffix');
      const hasRequiresTotal = columns.includes('requires_total');
      const hasIsPrimary = columns.includes('is_primary');
      if (!hasFileSuffix && !hasRequiresTotal && !hasIsPrimary) {
        if (typeof done === 'function') done();
        return;
      }

      console.log('Migrating document_definitions table to drop deprecated columns…');
      db.serialize(() => {
        db.run('BEGIN');
        db.run(
          `CREATE TABLE IF NOT EXISTS document_definitions_new (
            definition_id INTEGER PRIMARY KEY AUTOINCREMENT,
            business_id INTEGER NOT NULL,
            key TEXT NOT NULL,
            doc_type TEXT NOT NULL,
            label TEXT NOT NULL,
            description TEXT,
            invoice_variant TEXT,
            template_path TEXT,
            is_active INTEGER DEFAULT 1,
            is_locked INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            sheet_exports TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (business_id) REFERENCES business_settings(id),
            UNIQUE (business_id, key)
          )`
        );
        db.run(
          `INSERT INTO document_definitions_new (
             definition_id, business_id, key, doc_type, label, description, invoice_variant, template_path, is_active, is_locked, sort_order, sheet_exports, created_at, updated_at
           )
           SELECT definition_id, business_id, key, doc_type, label, description, invoice_variant, template_path, is_active, is_locked, sort_order, sheet_exports, created_at, updated_at
           FROM document_definitions`
        );
        db.run('DROP TABLE document_definitions');
        db.run('ALTER TABLE document_definitions_new RENAME TO document_definitions');
        db.run(`CREATE INDEX IF NOT EXISTS idx_document_definitions_business_sort ON document_definitions (business_id, sort_order)`);
        db.run('COMMIT', (commitErr) => {
          if (commitErr) {
            console.error('Failed to commit document_definitions migration', commitErr);
          }
          if (typeof done === 'function') done();
        });
      });
    });
  } catch (err) {
    console.error('Migration error for document_definitions', err);
    if (typeof done === 'function') done();
  }
}

// Drop legacy template path columns from business_settings after migrating their values into document_definitions
function migrateBusinessSettingsDropTemplateColumns(done) {
  try {
    db.all(`PRAGMA table_info(business_settings)`, (err, rows) => {
      if (err) { if (typeof done === 'function') done(); return; }
      const columns = Array.isArray(rows) ? rows.map(r => r.name) : [];
      const legacy = ['invoice_template_path', 'quote_template_path', 'contract_template_path', 'gig_sheet_template_path'];
      const hasAnyLegacy = legacy.some(c => columns.includes(c));
      if (!hasAnyLegacy) { if (typeof done === 'function') done(); return; }

      // Migrate any existing business-level template paths into definitions first
      db.all(
        `SELECT id, invoice_template_path, quote_template_path, contract_template_path FROM business_settings`,
        [],
        (selErr, businesses) => {
          if (!selErr && Array.isArray(businesses) && businesses.length) {
            businesses.forEach(biz => {
              const id = biz.id;
              const apply = (value, keys) => {
                if (!value || !keys || !keys.length) return;
                keys.forEach(key => {
                  db.run(
                    "UPDATE document_definitions SET template_path = COALESCE(template_path, ?) WHERE business_id = ? AND key = ?",
                    [value, id, key],
                    () => {}
                  );
                });
              };
              apply(biz.invoice_template_path || null, ['invoice_deposit', 'invoice_balance']);
              apply(biz.quote_template_path || null, ['quote']);
              apply(biz.contract_template_path || null, ['contract']);
            });
          }

          // Rebuild table without legacy columns
          db.serialize(() => {
            // Temporarily disable foreign key enforcement for schema change
            db.run('PRAGMA foreign_keys = OFF');
            db.run('BEGIN');
            // Ensure staging table is clean to avoid UNIQUE constraint issues on repeated runs
            db.run('DROP TABLE IF EXISTS business_settings_new');
            db.run(
              `CREATE TABLE IF NOT EXISTS business_settings_new (
                id INTEGER PRIMARY KEY,
                business_name TEXT NOT NULL UNIQUE,
                last_invoice_number INTEGER DEFAULT 0,
                last_quote_number INTEGER DEFAULT 0,
                save_path TEXT NOT NULL
              )`
            );
            db.run(
              `INSERT INTO business_settings_new (id, business_name, last_invoice_number, last_quote_number, save_path)
               SELECT id, business_name, last_invoice_number, last_quote_number, save_path FROM business_settings`
            );
            db.run('DROP TABLE business_settings');
            db.run('ALTER TABLE business_settings_new RENAME TO business_settings');
            db.run('COMMIT', () => {
              // Re-enable foreign key enforcement after migration completes
              db.run('PRAGMA foreign_keys = ON');
              if (typeof done === 'function') done();
            });
          });
        }
      );
    });
  } catch (_err) {
    if (typeof done === 'function') done();
  }
}

function syncLegacyBusinessesTable() {
  DEFAULT_BUSINESSES.forEach(business => {
    db.run(
      `INSERT OR IGNORE INTO businesses (business_id, name) VALUES (?, ?)`,
      [business.id, business.business_name]
    );

    db.run(
      `UPDATE businesses SET name = ? WHERE business_id = ?`,
      [business.business_name, business.id]
    );
  });
}

function mapDocumentDefinitionRow(row) {
  if (!row) return null;
  let sheetExports = [];
  if (row.sheet_exports) {
    try {
      const parsed = typeof row.sheet_exports === 'string' ? JSON.parse(row.sheet_exports) : row.sheet_exports;
      if (Array.isArray(parsed)) sheetExports = parsed;
    } catch (err) {
      console.warn('Unable to parse sheet_exports for definition', row.key, err);
    }
  }
  return {
    ...row,
    is_primary: 0,
    is_active: row.is_active ? 1 : 0,
    is_locked: row.is_locked ? 1 : 0,
    sort_order: Number.isFinite(row.sort_order) ? Number(row.sort_order) : 0,
    sheet_exports: sheetExports
  };
}

function getDocumentDefinitionsForBusiness(businessId, options = {}) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    const includeInactive = options.includeInactive === true;
    const whereParts = ['business_id = ?'];
    if (!includeInactive) {
      whereParts.push('is_active = 1');
    }

    db.all(
      `SELECT * FROM document_definitions WHERE ${whereParts.join(' AND ')} ORDER BY sort_order ASC, label COLLATE NOCASE ASC`,
      [id],
      (err, rows) => {
        if (err) reject(err);
        else resolve((rows || []).map(mapDocumentDefinitionRow));
      }
    );
  });
}

function getDocumentDefinitionRecord(businessId, identifier) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    if (identifier == null || identifier === '') {
      resolve(null);
      return;
    }

    const query = typeof identifier === 'number'
      ? 'SELECT * FROM document_definitions WHERE business_id = ? AND definition_id = ?'
      : 'SELECT * FROM document_definitions WHERE business_id = ? AND key = ?';

    db.get(query, [id, identifier], (err, row) => {
      if (err) reject(err);
      else resolve(mapDocumentDefinitionRow(row));
    });
  });
}

function determineNextDefinitionSortOrder(businessId) {
  return new Promise((resolve, reject) => {
    db.get(
      'SELECT COALESCE(MAX(sort_order), -1) AS max_order FROM document_definitions WHERE business_id = ?',
      [businessId],
      (err, row) => {
        if (err) reject(err);
        else resolve((Number(row?.max_order) || 0) + 1);
      }
    );
  });
}

function sanitizeDefinitionPayload(definition) {
  if (!definition || typeof definition !== 'object') return {};
  let sheetExportsValue = null;
  if (Array.isArray(definition.sheet_exports)) {
    sheetExportsValue = JSON.stringify(definition.sheet_exports);
  } else if (typeof definition.sheet_exports === 'string') {
    const trimmed = definition.sheet_exports.trim();
    sheetExportsValue = trimmed ? trimmed : null;
  }
  return {
    key: (definition.key || '').trim(),
    doc_type: (definition.doc_type || '').trim(),
    label: (definition.label || '').trim(),
    description: definition.description != null && definition.description !== '' ? String(definition.description) : null,
    invoice_variant: definition.invoice_variant != null && definition.invoice_variant !== '' ? String(definition.invoice_variant) : null,
    template_path: definition.template_path != null && definition.template_path !== '' ? String(definition.template_path) : null,
    is_active: definition.is_active === 0 ? 0 : 1,
    is_locked: definition.is_locked ? 1 : 0,
    sort_order: Number.isFinite(definition.sort_order) ? Number(definition.sort_order) : null,
    sheet_exports: sheetExportsValue
  };
}

function saveDocumentDefinition(businessId, definition) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    const payload = sanitizeDefinitionPayload(definition);
    if (!payload.key) {
      reject(new Error('Definition key is required'));
      return;
    }
    if (!payload.doc_type) {
      reject(new Error('Document type is required')); return;
    }
    if (!payload.label) {
      reject(new Error('Definition label is required'));
      return;
    }

    const proceed = (sortOrder) => {
      const now = new Date().toISOString();
      const orderValue = sortOrder != null ? sortOrder : 0;

      db.run(
        `INSERT INTO document_definitions (
           business_id,
           key,
           doc_type,
           label,
           description,
           invoice_variant,
           template_path,
           is_active,
           is_locked,
           sort_order,
           sheet_exports,
           created_at,
           updated_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(business_id, key) DO UPDATE SET
           doc_type = excluded.doc_type,
           label = excluded.label,
           description = excluded.description,
           invoice_variant = excluded.invoice_variant,
           template_path = excluded.template_path,
           is_active = excluded.is_active,
           sort_order = excluded.sort_order,
           sheet_exports = excluded.sheet_exports,
           updated_at = excluded.updated_at,
           is_locked = excluded.is_locked`,
        [
          id,
          payload.key,
          payload.doc_type,
          payload.label,
          payload.description,
          payload.invoice_variant,
          payload.template_path,
          payload.is_active,
          payload.is_locked,
          orderValue,
          payload.sheet_exports,
          now,
          now
        ],
        function (err) {
          if (err) {
            reject(err);
          } else {
            db.run(
              'DELETE FROM document_definition_tombstones WHERE business_id = ? AND key = ?',
              [id, payload.key],
              () => {}
            );
            resolve({ key: payload.key, changes: this.changes });
          }
        }
      );
    };

    if (payload.sort_order == null) {
      determineNextDefinitionSortOrder(id)
        .then(order => proceed(order))
        .catch(reject);
    } else {
      proceed(payload.sort_order);
    }
  });
}

function deleteDocumentDefinition(businessId, identifier) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    if (identifier == null || identifier === '') {
      reject(new Error('Definition identifier is required'));
      return;
    }

    const query = typeof identifier === 'number'
      ? 'SELECT definition_id, key, is_locked FROM document_definitions WHERE business_id = ? AND definition_id = ?'
      : 'SELECT definition_id, key, is_locked FROM document_definitions WHERE business_id = ? AND key = ?';

    db.get(query, [id, identifier], (err, row) => {
      if (err) {
        reject(err);
        return;
      }
      if (!row) {
        resolve({ removed: 0 });
        return;
      }
      if (row.is_locked) {
        reject(new Error('Locked document definitions cannot be deleted'));
        return;
      }

      const resolvedKey = row.key;
      if (resolvedKey) {
        db.run(
          'INSERT OR IGNORE INTO document_definition_tombstones (business_id, key) VALUES (?, ?)',
          [id, resolvedKey],
          () => {}
        );
      }

      db.run(
        "DELETE FROM document_definitions WHERE definition_id = ?",
        [row.definition_id],
        function (deleteErr) {
          if (deleteErr) reject(deleteErr);
          else resolve({ removed: this.changes });
        }
      );
    });
  });
}

function deleteDocumentsByDefinition(businessId, definitionKey, options = {}) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    const key = (definitionKey || '').trim();
    if (!key) {
      reject(new Error('definition_key is required'));
      return;
    }

    const conditions = ['business_id = ?', 'definition_key = ?'];
    const params = [id, key];

    if (options.docType) {
      conditions.push('doc_type = ?');
      params.push(options.docType);
    }

    if (options.jobsheetId !== undefined) {
      if (options.jobsheetId === null) {
        conditions.push('jobsheet_id IS NULL');
      } else {
        conditions.push('jobsheet_id = ?');
        params.push(Number(options.jobsheetId));
      }
    }

    if (options.eventId !== undefined) {
      if (options.eventId === null) {
        conditions.push('event_id IS NULL');
      } else {
        conditions.push('event_id = ?');
        params.push(Number(options.eventId));
      }
    }

    const selectParams = params.slice();
    const deleteParams = params.slice();

    db.all(
      `SELECT document_id, file_path FROM documents WHERE ${conditions.join(' AND ')}`,
      selectParams,
      (selectErr, rows) => {
        if (selectErr) {
          reject(selectErr);
          return;
        }

        db.run(
          `DELETE FROM documents WHERE ${conditions.join(' AND ')}`,
          deleteParams,
          function (err) {
            if (err) reject(err);
            else resolve({ removed: this.changes || 0, documents: rows || [] });
          }
        );
      }
    );
  });
}

function deleteDocumentsByPathPrefix(businessId, absolutePath) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    if (!absolutePath) {
      resolve({ removed: 0 });
      return;
    }

    const normalized = path.resolve(absolutePath);
    const escaped = escapeLikePattern(normalized);
    const likeValue = `${escaped}%`;

    db.run(
      "DELETE FROM documents WHERE business_id = ? AND file_path LIKE ? ESCAPE '\\'",
      [id, likeValue],
      function (err) {
        if (err) reject(err);
        else resolve({ removed: this.changes || 0 });
      }
    );
  });
}

function deleteDocumentByFilePath(businessId, absolutePath) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    const normalized = absolutePath ? path.resolve(absolutePath) : null;
    if (!normalized) {
      resolve({ removed: 0 });
      return;
    }

    db.run(
      'DELETE FROM documents WHERE business_id = ? AND file_path = ?',
      [id, normalized],
      function (err) {
        if (err) reject(err);
        else resolve({ removed: this.changes || 0 });
      }
    );
  });
}

function fetchBusinessRecord(businessId) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }
    db.get(
      'SELECT * FROM business_settings WHERE id = ?',
      [id],
      (err, row) => {
        if (err) reject(err);
        else resolve(row || null);
      }
    );
  });
}

async function pathExists(targetPath) {
  if (!targetPath) return false;
  try {
    await fsp.access(targetPath);
    return true;
  } catch (err) {
    return false;
  }
}

function resolveRelativePath(rootPath, relativePath) {
  if (!relativePath) {
    return path.resolve(rootPath);
  }
  const segments = relativePath.split(/[\\/]+/).filter(Boolean);
  return path.resolve(path.join(rootPath, ...segments));
}

function isSubPath(parentPath, childPath) {
  const parentResolved = path.resolve(parentPath);
  const childResolved = path.resolve(childPath);
  const relative = path.relative(parentResolved, childResolved);
  return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative));
}

async function ensureDirectoryExists(targetPath) {
  if (!targetPath) return;
  await fsp.mkdir(targetPath, { recursive: true });
}

function appendSuffix(name, counter) {
  if (!counter) return name;
  const ext = path.extname(name);
  const base = ext ? name.slice(0, -ext.length) : name;
  return `${base} (${counter})${ext}`;
}

async function ensureUniquePath(directory, name) {
  const baseName = name || 'item';
  let attempt = 0;
  while (attempt < 1000) {
    const candidateName = attempt === 0 ? baseName : appendSuffix(baseName, attempt);
    const candidatePath = path.join(directory, candidateName);
    if (!(await pathExists(candidatePath))) {
      return candidatePath;
    }
    attempt += 1;
  }
  return path.join(directory, `${baseName}-${Date.now()}`);
}

async function copyRecursive(source, destination) {
  const stats = await fsp.lstat(source);
  if (stats.isSymbolicLink()) {
    try {
      const target = await fsp.readlink(source);
      await ensureDirectoryExists(path.dirname(destination));
      await fsp.symlink(target, destination);
    } catch (err) {
      await ensureDirectoryExists(path.dirname(destination));
      await fsp.copyFile(source, destination);
    }
    return;
  }

  if (stats.isDirectory()) {
    await fsp.mkdir(destination, { recursive: true });
    const entries = await fsp.readdir(source);
    for (const entry of entries) {
      const fromChild = path.join(source, entry);
      const toChild = path.join(destination, entry);
      await copyRecursive(fromChild, toChild);
    }
    return;
  }

  await ensureDirectoryExists(path.dirname(destination));
  await fsp.copyFile(source, destination);
}

async function removeRecursive(targetPath) {
  await fsp.rm(targetPath, { recursive: true, force: true });
}

async function moveEntryToTrash(rootPath, targetPath) {
  const rootResolved = path.resolve(rootPath);
  const targetResolved = path.resolve(targetPath);
  if (!isSubPath(rootResolved, targetResolved) || targetResolved === rootResolved) {
    throw new Error('Target must be within the documents folder.');
  }

  const trashRoot = path.join(rootResolved, TRASH_DIR_NAME);
  await ensureDirectoryExists(trashRoot);
  const destination = await ensureUniquePath(trashRoot, path.basename(targetResolved) || 'item');

  try {
    await fsp.rename(targetResolved, destination);
  } catch (err) {
    if (err && err.code === 'EXDEV') {
      await copyRecursive(targetResolved, destination);
      await removeRecursive(targetResolved);
    } else if (err && err.code === 'ENOENT') {
      throw new Error('Path not found.');
    } else {
      throw err;
    }
  }

  return destination;
}

async function summarizeDirectory(targetPath) {
  const stats = await fsp.lstat(targetPath);
  if (!stats.isDirectory()) {
    return {
      itemCount: 1,
      totalSize: stats.size || 0
    };
  }

  const entries = await fsp.readdir(targetPath);
  let totalSize = 0;
  let itemCount = 0;
  for (const entry of entries) {
    const childPath = path.join(targetPath, entry);
    try {
      const summary = await summarizeDirectory(childPath);
      totalSize += summary.totalSize;
      itemCount += summary.itemCount;
    } catch (err) {
      // ignore inaccessible entries
    }
  }
  return { itemCount, totalSize };
}

async function buildDirectoryTree(currentPath, relativePath = '', depth = 0) {
  const stats = await fsp.lstat(currentPath);
  const isDirectory = stats.isDirectory();
  const baseName = path.basename(currentPath) || (relativePath ? relativePath.split('/').pop() : 'Documents');
  const node = {
    name: baseName || 'Documents',
    path: relativePath,
    absolutePath: currentPath,
    isDirectory,
    size: stats.size || 0,
    modified: stats.mtime ? stats.mtime.toISOString() : null
  };

  if (!isDirectory || depth >= MAX_TREE_DEPTH) {
    node.children = [];
    node.itemCount = isDirectory ? 0 : 1;
    node.totalSize = stats.size || 0;
    return node;
  }

  let entries = [];
  try {
    entries = await fsp.readdir(currentPath, { withFileTypes: true });
  } catch (err) {
    node.children = [];
    node.itemCount = 0;
    node.totalSize = stats.size || 0;
    return node;
  }

  const children = [];
  let totalSize = 0;
  let totalItems = 0;
  let processed = 0;

  for (const entry of entries) {
    if (entry.name === '.' || entry.name === '..') continue;
    if (entry.name === TRASH_DIR_NAME) continue;
    if (processed >= MAX_TREE_ENTRIES) break;
    processed += 1;

    const childAbsolute = path.join(currentPath, entry.name);
    const childRelative = relativePath ? `${relativePath}/${entry.name}` : entry.name;

    try {
      const childNode = await buildDirectoryTree(childAbsolute, childRelative, depth + 1);
      children.push(childNode);
      totalSize += childNode.totalSize != null ? childNode.totalSize : (childNode.size || 0);
      totalItems += childNode.isDirectory ? (childNode.itemCount || 0) : 1;
    } catch (err) {
      // skip problematic entries
    }
  }

  children.sort((a, b) => {
    if (a.isDirectory && !b.isDirectory) return -1;
    if (!a.isDirectory && b.isDirectory) return 1;
    return a.name.localeCompare(b.name, 'en', { sensitivity: 'base' });
  });

  node.children = children;
  node.itemCount = totalItems;
  node.totalSize = totalSize;
  node.size = totalSize;
  return node;
}

async function summarizeTrashDirectory(trashPath) {
  const exists = await pathExists(trashPath);
  if (!exists) {
    return {
      path: TRASH_DIR_NAME,
      absolutePath: trashPath,
      itemCount: 0,
      size: 0
    };
  }

  const summary = await summarizeDirectory(trashPath);
  return {
    path: TRASH_DIR_NAME,
    absolutePath: trashPath,
    itemCount: summary.itemCount,
    size: summary.totalSize
  };
}

function setDocumentFilePath(documentId, filePath) {
  return new Promise((resolve, reject) => {
    const id = Number(documentId);
    if (!Number.isInteger(id)) {
      resolve();
      return;
    }
    db.run(
      `UPDATE documents SET file_path = ?, updated_at = datetime('now') WHERE document_id = ?`,
      [filePath || null, id],
      function (err) {
        if (err) reject(err);
        else resolve();
      }
    );
  });
}

async function updateDocumentPathsUnder(businessId, sourceBase, targetBase) {
  const id = Number(businessId);
  if (!Number.isInteger(id)) return;
  if (!sourceBase || !targetBase) return;

  const normalizedSource = path.resolve(sourceBase);
  const normalizedTarget = path.resolve(targetBase);
  const escaped = escapeLikePattern(normalizedSource);
  const likeValue = `${escaped}%`;

  const rows = await new Promise((resolve, reject) => {
    db.all(
      "SELECT document_id, file_path FROM documents WHERE business_id = ? AND file_path LIKE ? ESCAPE '\\'",
      [id, likeValue],
      (err, data) => {
        if (err) reject(err);
        else resolve(data || []);
      }
    );
  });

  await Promise.all(rows.map(row => {
    if (!row.file_path || typeof row.file_path !== 'string') return null;
    const remainder = row.file_path.slice(normalizedSource.length).replace(/^[\\/]+/, '');
    const nextPath = path.join(normalizedTarget, remainder);
    return setDocumentFilePath(row.document_id, nextPath);
  }));
}

async function clearDocumentPath(businessId, absolutePath) {
  const id = Number(businessId);
  if (!Number.isInteger(id) || !absolutePath) return;
  const normalized = path.resolve(absolutePath);
  await new Promise((resolve, reject) => {
    db.run(
      `UPDATE documents SET file_path = NULL, updated_at = datetime('now') WHERE business_id = ? AND file_path = ?`,
      [id, normalized],
      function (err) {
        if (err) reject(err);
        else resolve();
      }
    );
  });
}

async function clearDocumentPathsByPrefix(businessId, absolutePath) {
  const id = Number(businessId);
  if (!Number.isInteger(id) || !absolutePath) return;
  const normalized = path.resolve(absolutePath);
  const escaped = escapeLikePattern(normalized);
  const likeValue = `${escaped}%`;
  await new Promise((resolve, reject) => {
    db.run(
      "UPDATE documents SET file_path = NULL, updated_at = datetime('now') WHERE business_id = ? AND file_path LIKE ? ESCAPE '\\'",
      [id, likeValue],
      function (err) {
        if (err) reject(err);
        else resolve();
      }
    );
  });
}

async function relocateBusinessDocuments(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required');
  }

  const targetPathRaw = options.targetPath || options.destinationPath;
  if (!targetPathRaw) {
    throw new Error('targetPath is required');
  }
  const targetPath = path.resolve(targetPathRaw);
  await ensureDirectoryExists(targetPath);

  const summary = { moved: [], skipped: [], errors: [] };
  const sourcePathRaw = options.sourcePath || options.originPath || null;
  if (!sourcePathRaw) {
    await ensureDirectoryExists(path.join(targetPath, TRASH_DIR_NAME));
    return summary;
  }

  const sourcePath = path.resolve(sourcePathRaw);
  if (sourcePath === targetPath) {
    await ensureDirectoryExists(path.join(targetPath, TRASH_DIR_NAME));
    return summary;
  }

  if (!(await pathExists(sourcePath))) {
    await ensureDirectoryExists(path.join(targetPath, TRASH_DIR_NAME));
    return summary;
  }

  const entries = await fsp.readdir(sourcePath);
  for (const entry of entries) {
    const fromPath = path.join(sourcePath, entry);
    const toPath = path.join(targetPath, entry);
    if (await pathExists(toPath)) {
      summary.skipped.push({ name: entry, reason: 'exists' });
      continue;
    }

    try {
      await ensureDirectoryExists(path.dirname(toPath));
      await fsp.rename(fromPath, toPath);
      summary.moved.push({ from: fromPath, to: toPath });
      await updateDocumentPathsUnder(businessId, fromPath, toPath);
    } catch (err) {
      if (err && err.code === 'EXDEV') {
        try {
          await copyRecursive(fromPath, toPath);
          await removeRecursive(fromPath);
          summary.moved.push({ from: fromPath, to: toPath, copied: true });
          await updateDocumentPathsUnder(businessId, fromPath, toPath);
        } catch (copyErr) {
          summary.errors.push({ name: entry, message: copyErr.message || String(copyErr) });
        }
      } else {
        summary.errors.push({ name: entry, message: err?.message || String(err) });
      }
    }
  }

  await ensureDirectoryExists(path.join(targetPath, TRASH_DIR_NAME));
  return summary;
}

async function listDocumentTree(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required');
  }

  const business = await fetchBusinessRecord(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured');
  }

  const rootPath = path.resolve(business.save_path);
  await ensureDirectoryExists(rootPath);
  const rootNode = await buildDirectoryTree(rootPath);
  const trashSummary = await summarizeTrashDirectory(path.join(rootPath, TRASH_DIR_NAME));

  return {
    rootPath,
    root: rootNode,
    trash: trashSummary
  };
}

async function deleteDocumentFolder(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required');
  }

  const relativePath = options.relativePath || options.path;
  if (!relativePath) {
    throw new Error('relativePath is required');
  }

  const business = await fetchBusinessRecord(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured');
  }

  const rootPath = path.resolve(business.save_path);
  const targetPath = resolveRelativePath(rootPath, relativePath);
  if (!(await pathExists(targetPath))) {
    throw new Error('Folder not found');
  }
  const stats = await fsp.lstat(targetPath);
  if (!stats.isDirectory()) {
    throw new Error('Target is not a folder');
  }

  // Prevent deleting folders that contain locked documents
  const escaped = escapeLikePattern(path.resolve(targetPath));
  const likeValue = `${escaped}%`;
  const lockedCount = await new Promise((resolve, reject) => {
    db.get(
      "SELECT COUNT(1) AS c FROM documents WHERE business_id = ? AND file_path LIKE ? ESCAPE '\\' AND is_locked = 1",
      [businessId, likeValue],
      (err, row) => {
        if (err) reject(err);
        else resolve(row?.c || 0);
      }
    );
  });
  if (lockedCount > 0) {
    throw new Error('Cannot delete folder: contains locked documents');
  }

  const trashedPath = await moveEntryToTrash(rootPath, targetPath);
  await clearDocumentPathsByPrefix(businessId, targetPath);
  return { ok: true, trashedPath };
}

async function deleteDocumentByPath(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required');
  }

  const absolutePathRaw = options.absolutePath || options.path;
  if (!absolutePathRaw) {
    throw new Error('absolutePath is required');
  }

  const business = await fetchBusinessRecord(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured');
  }

  const rootPath = path.resolve(business.save_path);
  const targetPath = path.resolve(absolutePathRaw);
  if (!isSubPath(rootPath, targetPath)) {
    throw new Error('Path must be inside the documents folder');
  }
  if (!(await pathExists(targetPath))) {
    throw new Error('File not found');
  }

  const stats = await fsp.lstat(targetPath);
  if (stats.isDirectory()) {
    const relative = path.relative(rootPath, targetPath);
    return deleteDocumentFolder({ businessId, relativePath: relative });
  }

  // Block deletion of locked documents
  const lockedRow = await new Promise((resolve, reject) => {
    db.get(
      `SELECT is_locked FROM documents WHERE business_id = ? AND file_path = ? LIMIT 1`,
      [businessId, targetPath],
      (err, row) => {
        if (err) reject(err);
        else resolve(row || null);
      }
    );
  });
  if (lockedRow && lockedRow.is_locked) {
    throw new Error('This document is locked and cannot be deleted');
  }

  const trashedPath = await moveEntryToTrash(rootPath, targetPath);
  await clearDocumentPath(businessId, targetPath);
  return { ok: true, trashedPath };
}

async function emptyDocumentsTrash(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  if (!Number.isInteger(businessId)) {
    throw new Error('businessId is required');
  }

  const business = await fetchBusinessRecord(businessId);
  if (!business || !business.save_path) {
    throw new Error('Documents folder not configured');
  }

  const rootPath = path.resolve(business.save_path);
  const trashPath = path.join(rootPath, TRASH_DIR_NAME);
  if (!(await pathExists(trashPath))) {
    await ensureDirectoryExists(trashPath);
    return { ok: true, removed: 0 };
  }

  const summary = await summarizeTrashDirectory(trashPath);
  await removeRecursive(trashPath);
  await ensureDirectoryExists(trashPath);
  return { ok: true, removed: summary?.itemCount || 0 };
}

function reorderDocumentDefinitions(businessId, orderedKeys) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }
    if (!Array.isArray(orderedKeys)) {
      resolve({ changes: 0 });
      return;
    }

    const stmt = db.prepare("UPDATE document_definitions SET sort_order = ?, updated_at = datetime('now') WHERE business_id = ? AND key = ?");
    orderedKeys.forEach((key, index) => {
      if (!key) return;
      stmt.run(index, id, key);
    });
    stmt.finalize(err => {
      if (err) reject(err);
      else resolve({ changes: orderedKeys.length });
    });
  });
}

function getJobsheetTemplateOverrides(jobsheetId) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }

    db.all(
      `SELECT definition_key, template_path FROM jobsheet_template_overrides WHERE jobsheet_id = ?`,
      [id],
      (err, rows) => {
        if (err) reject(err);
        else resolve(rows || []);
      }
    );
  });
}

function setJobsheetTemplateOverride(jobsheetId, definitionKey, templatePath) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }
    const key = (definitionKey || '').trim();
    if (!key) {
      reject(new Error('Definition key is required'));
      return;
    }
    const pathValue = (templatePath || '').trim();
    if (!pathValue) {
      reject(new Error('Template path is required'));
      return;
    }

    db.run(
      `INSERT INTO jobsheet_template_overrides (jobsheet_id, definition_key, template_path, updated_at)
       VALUES (?, ?, ?, datetime('now'))
       ON CONFLICT(jobsheet_id, definition_key) DO UPDATE SET
         template_path = excluded.template_path,
         updated_at = excluded.updated_at`,
      [id, key, pathValue],
      function (err) {
        if (err) reject(err);
        else resolve({ overrideId: this.lastID || null, template_path: pathValue });
      }
    );
  });
}

function clearJobsheetTemplateOverride(jobsheetId, definitionKey) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }
    const key = (definitionKey || '').trim();
    if (!key) {
      reject(new Error('Definition key is required'));
      return;
    }

    db.run(
      `DELETE FROM jobsheet_template_overrides WHERE jobsheet_id = ? AND definition_key = ?`,
      [id, key],
      function (err) {
        if (err) reject(err);
        else resolve({ removed: this.changes || 0 });
      }
    );
  });
}

// No longer sync template paths from business settings; managed per-definition only

function getMergeFields() {
  return new Promise((resolve, reject) => {
    db.all(
      `SELECT
         f.field_key,
         f.label,
         f.placeholder,
         f.category,
         f.description,
         f.show_in_jobsheet,
         f.active,
         f.created_at,
         f.updated_at,
         b.template,
         b.sheet,
         b.cell,
         b.data_type,
         b.style,
         b.format
       FROM ${MERGE_FIELD_TABLE} f
       LEFT JOIN ${MERGE_FIELD_BINDINGS_TABLE} b ON b.field_key = f.field_key
       ORDER BY f.field_key`,
      (err, rows) => {
        if (err) return reject(err);
        const map = new Map();
        const result = [];
        (rows || []).forEach(row => {
          let entry = map.get(row.field_key);
          if (!entry) {
            entry = {
              field_key: row.field_key,
              label: row.label,
              placeholder: row.placeholder,
              category: row.category,
              description: row.description,
              show_in_jobsheet: row.show_in_jobsheet ? true : false,
              active: row.active ? true : false,
              created_at: row.created_at,
              updated_at: row.updated_at,
              bindings: []
            };
            map.set(row.field_key, entry);
            result.push(entry);
          }
          if (row.template) {
            entry.bindings.push({
              template: row.template,
              sheet: row.sheet,
              cell: row.cell,
              data_type: row.data_type,
              style: row.style,
              format: row.format
            });
          }
        });
        resolve(result);
      }
    );
  });
}

function getMergeFieldBindingsByTemplate(template) {
  return new Promise((resolve, reject) => {
    db.all(
      `SELECT b.field_key, b.template, b.sheet, b.cell, b.data_type, b.style, b.format, f.placeholder
       FROM ${MERGE_FIELD_BINDINGS_TABLE} b
       JOIN ${MERGE_FIELD_TABLE} f ON f.field_key = b.field_key
       WHERE b.template = ? AND f.active = 1`,
      [template],
      (err, rows) => {
        if (err) return reject(err);
        resolve(rows || []);
      }
    );
  });
}

function getMergeFieldValueSources(fieldKeys) {
  return new Promise((resolve, reject) => {
    let sql = `SELECT field_key, source_type, source_path, literal_value FROM merge_field_value_sources`;
    let params = [];

    if (Array.isArray(fieldKeys) && fieldKeys.length) {
      const placeholders = fieldKeys.map(() => '?').join(', ');
      sql += ` WHERE field_key IN (${placeholders})`;
      params = fieldKeys;
    }

    db.all(sql, params, (err, rows) => {
      if (err) {
        reject(err);
        return;
      }
      const map = {};
      (rows || []).forEach(row => {
        map[row.field_key] = {
          field_key: row.field_key,
          source_type: row.source_type,
          source_path: row.source_path || null,
          literal_value: row.literal_value || null
        };
      });
      resolve(map);
    });
  });
}

function setMergeFieldValueSource(fieldKey, source) {
  return new Promise((resolve, reject) => {
    if (!fieldKey || typeof fieldKey !== 'string') {
      reject(new Error('field_key is required'));
      return;
    }

    const sourceType = source?.source_type || source?.sourceType || null;
    const sourcePath = source?.source_path || source?.sourcePath || null;
    const literalValue = source?.literal_value || source?.literalValue || null;

    if (!sourceType) {
      reject(new Error('source_type is required'));
      return;
    }

    if (sourceType === 'contextPath' && !sourcePath) {
      reject(new Error('source_path is required for contextPath sources'));
      return;
    }

    db.run(
      `INSERT INTO merge_field_value_sources (field_key, source_type, source_path, literal_value, created_at, updated_at)
       VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
       ON CONFLICT(field_key) DO UPDATE SET
         source_type = excluded.source_type,
         source_path = excluded.source_path,
         literal_value = excluded.literal_value,
         updated_at = datetime('now')`,
      [fieldKey, sourceType, sourcePath, literalValue],
      function (err) {
        if (err) reject(err);
        else resolve({ field_key: fieldKey });
      }
    );
  });
}

function clearMergeFieldValueSource(fieldKey) {
  return new Promise((resolve, reject) => {
    if (!fieldKey || typeof fieldKey !== 'string') {
      reject(new Error('field_key is required'));
      return;
    }

    db.run(
      `DELETE FROM merge_field_value_sources WHERE field_key = ?`,
      [fieldKey],
      function (err) {
        if (err) reject(err);
        else resolve({ removed: this.changes || 0 });
      }
    );
  });
}

function saveMergeField(field) {
  return new Promise((resolve, reject) => {
    if (!field || typeof field !== 'object') {
      reject(new Error('Field payload required'));
      return;
    }

    const key = field.field_key || field.fieldKey;
    if (!key || typeof key !== 'string') {
      reject(new Error('field_key is required'));
      return;
    }

    const label = field.label || key;
    const placeholder = field.placeholder || null;
    const category = field.category || null;
    const description = field.description || null;
    const showInJobsheet = field.show_in_jobsheet ?? field.showInJobsheet;
    const active = field.active == null ? true : Boolean(field.active);
    const bindings = Array.isArray(field.bindings) ? field.bindings : [];

    db.serialize(() => {
      db.run(
        `INSERT INTO ${MERGE_FIELD_TABLE} (field_key, label, placeholder, category, description, show_in_jobsheet, active, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
         ON CONFLICT(field_key) DO UPDATE SET
           label = excluded.label,
           placeholder = excluded.placeholder,
           category = excluded.category,
           description = excluded.description,
           show_in_jobsheet = excluded.show_in_jobsheet,
           active = excluded.active,
           updated_at = datetime('now')`,
        [
          key,
          label,
          placeholder,
          category,
          description,
          showInJobsheet ? 1 : 0,
          active ? 1 : 0
        ],
        err => {
          if (err) {
            reject(err);
            return;
          }

          db.run(
            `DELETE FROM ${MERGE_FIELD_BINDINGS_TABLE} WHERE field_key = ?`,
            [key],
            deleteErr => {
              if (deleteErr) {
                reject(deleteErr);
                return;
              }

              if (!bindings.length) {
                resolve({ field_key: key });
                return;
              }

              const insertBinding = db.prepare(
                `INSERT INTO ${MERGE_FIELD_BINDINGS_TABLE} (field_key, template, sheet, cell, data_type, style, format)
                 VALUES (?, ?, ?, ?, ?, ?, ?)`
              );

              bindings.forEach(binding => {
                if (!binding || typeof binding !== 'object') return;
                const template = binding.template;
                if (!template) return;
                insertBinding.run(
                  key,
                  template,
                  binding.sheet || null,
                  binding.cell || null,
                  binding.data_type || binding.dataType || 'string',
                  binding.style || null,
                  binding.format || null
                );
              });

              insertBinding.finalize(finalizeErr => {
                if (finalizeErr) {
                  reject(finalizeErr);
                  return;
                }
                resolve({ field_key: key });
              });
            }
          );
        }
      );
    });
  });
}

function deleteMergeField(fieldKey) {
  return new Promise((resolve, reject) => {
    if (!fieldKey) {
      reject(new Error('field_key is required'));
      return;
    }

    db.run(
      `DELETE FROM ${MERGE_FIELD_TABLE} WHERE field_key = ?`,
      [fieldKey],
      function (err) {
        if (err) {
          reject(err);
        } else {
          resolve({ deleted: this.changes });
        }
      }
    );
  });
}

initializeDatabase();

function updateDocumentTimestamp(documentId) {
  db.run(
    `UPDATE documents SET updated_at = datetime('now') WHERE document_id = ?`,
    [documentId]
  );
}

function updateBusinessSettingsRecord(businessId, updates = {}) {
  return new Promise((resolve, reject) => {
    const id = Number(businessId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid business id'));
      return;
    }

    if (!updates || typeof updates !== 'object') {
      resolve({ changes: 0, record: null });
      return;
    }

    const entries = Object.entries(updates).filter(([key]) => BUSINESS_SETTINGS_MUTABLE_FIELDS.has(key));
    if (!entries.length) {
      resolve({ changes: 0, record: null });
      return;
    }

    const setClauses = entries.map(([key]) => `${key} = ?`).join(', ');
    const values = entries.map(([, value]) => (value === undefined ? null : value));

    db.run(
      `UPDATE business_settings SET ${setClauses} WHERE id = ?`,
      [...values, id],
      function (err) {
        if (err) {
          reject(err);
          return;
        }

        const changes = this.changes;
        if (!changes) {
          resolve({ changes: 0, record: null });
          return;
        }

        db.get(
          `SELECT * FROM business_settings WHERE id = ?`,
          [id],
          (selectErr, row) => {
            if (selectErr) {
              reject(selectErr);
            } else {
              // No per-business template path syncing; definitions own template paths
              resolve({ changes, record: row || null });
            }
          }
        );
      }
    );
  });
}

function getCounterColumn(docType) {
  if (!docType) return null;
  const normalized = docType.toLowerCase();
  if (normalized === 'invoice') return 'last_invoice_number';
  if (normalized === 'quote') return 'last_quote_number';
  return null;
}

function sanitizeAhmenJobsheetValue(field, value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    value = trimmed;
  }
  if (AHMEN_JOBSHEET_BOOLEAN_FIELDS.has(field)) {
    if (value === true || value === 'true' || value === '1' || value === 1 || value === 'on') return 1;
    return 0;
  }
  if (AHMEN_JOBSHEET_INTEGER_FIELDS.has(field)) {
    const intVal = parseInt(value, 10);
    return Number.isInteger(intVal) ? intVal : null;
  }
  if (AHMEN_JOBSHEET_NUMERIC_FIELDS.has(field)) {
    const num = Number(value);
    return Number.isFinite(num) ? num : null;
  }
  return value;
}

function buildAhmenJobsheetValues(data) {
  return AHMEN_JOBSHEET_FIELDS.map(field => sanitizeAhmenJobsheetValue(field, data?.[field]));
}

function normalizeAhmenStatus(value) {
  if (typeof value !== 'string') return 'enquiry';
  const normalized = value.trim().toLowerCase();
  return AHMEN_JOBSHEET_STATUS_VALUES.has(normalized) ? normalized : 'enquiry';
}

function mapAhmenJobsheetRow(row) {
  if (!row) return row;
  const mapped = { ...row };
  AHMEN_JOBSHEET_FIELDS.forEach(field => {
    if (AHMEN_JOBSHEET_NUMERIC_FIELDS.has(field) && mapped[field] !== null && mapped[field] !== undefined) {
      mapped[field] = Number(mapped[field]);
    }
    if (AHMEN_JOBSHEET_BOOLEAN_FIELDS.has(field)) {
      mapped[field] = mapped[field] ? 1 : 0;
    }
    if (AHMEN_JOBSHEET_INTEGER_FIELDS.has(field) && mapped[field] !== null && mapped[field] !== undefined) {
      mapped[field] = Number(mapped[field]);
    }
  });
  return mapped;
}

function sanitizeAhmenVenueValue(field, value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    value = trimmed;
  }
  if (field === 'is_private') {
    return value === true || value === 'true' || value === '1' || value === 1 || value === 'on' ? 1 : 0;
  }
  return value;
}

function mapAhmenVenueRow(row) {
  if (!row) return row;
  return {
    ...row,
    is_private: row.is_private ? 1 : 0
  };
}

function getAhmenJobsheets(options = {}) {
  const conditions = [];
  const params = [];
  if (options.businessId) {
    conditions.push('business_id = ?');
    params.push(options.businessId);
  }
  if (!options.includeArchived) {
    conditions.push('archived_at IS NULL');
  }
  let query = 'SELECT * FROM ahmen_jobsheets';
  if (conditions.length) {
    query += ` WHERE ${conditions.join(' AND ')}`;
  }
  query += ' ORDER BY datetime(updated_at) DESC, datetime(created_at) DESC';

  return new Promise((resolve, reject) => {
    db.all(query, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows.map(mapAhmenJobsheetRow));
    });
  });
}

function getAhmenJobsheet(jobsheetId) {
  const id = Number(jobsheetId);
  if (!Number.isInteger(id)) {
    return Promise.reject(new Error('Invalid jobsheet id'));
  }
  return new Promise((resolve, reject) => {
    db.get('SELECT * FROM ahmen_jobsheets WHERE jobsheet_id = ?', [id], (err, row) => {
      if (err) reject(err);
      else resolve(mapAhmenJobsheetRow(row));
    });
  });
}

function addAhmenJobsheet(data) {
  return new Promise((resolve, reject) => {
    const businessId = Number(data?.business_id);
    if (!Number.isInteger(businessId)) {
      reject(new Error('Business id is required for AhMen jobsheet'));
      return;
    }

    const clientName = (data?.client_name || '').trim();
    if (!clientName) {
      reject(new Error('Client name is required'));
      return;
    }

    const values = buildAhmenJobsheetValues({ ...data, business_id: businessId });
    const placeholders = AHMEN_JOBSHEET_FIELDS.map(() => '?').join(', ');

    db.run(
      `INSERT INTO ahmen_jobsheets (${AHMEN_JOBSHEET_FIELDS.join(', ')}) VALUES (${placeholders})`,
      values,
      function (err) {
        if (err) reject(err);
        else resolve(this.lastID);
      }
    );
  });
}

function updateAhmenJobsheet(jobsheetId, data) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }

    const clientName = (data?.client_name || '').trim();
    if (!clientName) {
      reject(new Error('Client name is required'));
      return;
    }

    const values = buildAhmenJobsheetValues({ ...data });
    const setClause = AHMEN_JOBSHEET_FIELDS.map(field => `${field} = ?`).join(', ');

    db.run(
      `UPDATE ahmen_jobsheets SET ${setClause}, updated_at = datetime('now') WHERE jobsheet_id = ?`,
      [...values, id],
      function (err) {
        if (err) reject(err);
        else resolve(this.changes);
      }
    );
  });
}

function updateAhmenJobsheetStatus(jobsheetId, status) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }

    const normalizedStatus = normalizeAhmenStatus(status);

    db.run(
      "UPDATE ahmen_jobsheets SET status = ?, updated_at = datetime('now') WHERE jobsheet_id = ?",
      [normalizedStatus, id],
      function (err) {
        if (err) reject(err);
        else resolve(this.changes);
      }
    );
  });
}

function setJobsheetArchived(jobsheetId, archived) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }
    const valueExpr = archived ? "datetime('now')" : 'NULL';
    db.run(
      `UPDATE ahmen_jobsheets SET archived_at = ${valueExpr}, updated_at = datetime('now') WHERE jobsheet_id = ?`,
      [id],
      function (err) {
        if (err) reject(err);
        else resolve(this.changes);
      }
    );
  });
}

function deleteAhmenJobsheet(jobsheetId) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }

    db.serialize(() => {
      db.run('DELETE FROM planner_actions WHERE jobsheet_id = ?', [id], () => {});
      db.run(
        'DELETE FROM ahmen_jobsheets WHERE jobsheet_id = ?',
        [id],
        function (err) {
          if (err) reject(err);
          else resolve(this.changes);
        }
      );
    });
  });
}

async function deleteJobsheetCompletely(options = {}) {
  const businessId = Number(options.businessId ?? options.business_id ?? options.id);
  const jobsheetId = Number(options.jobsheetId ?? options.jobsheet_id);
  const removeFiles = options.removeFiles !== false; // default true
  if (!Number.isInteger(businessId)) throw new Error('businessId is required');
  if (!Number.isInteger(jobsheetId)) throw new Error('jobsheetId is required');

  // Gather document paths first (by jobsheet id only to avoid FK stragglers where business_id may not match)
  const docs = await new Promise((resolve, reject) => {
    db.all(
      `SELECT document_id, file_path FROM documents WHERE jobsheet_id = ?`,
      [jobsheetId],
      (err, rows) => {
        if (err) reject(err);
        else resolve(Array.isArray(rows) ? rows : []);
      }
    );
  });

  const results = { filesTrashed: 0, documentsRemoved: 0, emailsRemoved: 0, scheduledRemoved: 0, plannerRemoved: 0, overridesRemoved: 0, jobsheetsRemoved: 0, fileErrors: [] };

  if (removeFiles && docs.length) {
    for (const row of docs) {
      const p = row?.file_path ? String(row.file_path) : '';
      if (!p) continue;
      try {
        await deleteDocumentByPath({ businessId, absolutePath: p });
        results.filesTrashed += 1;
      } catch (err) {
        results.fileErrors.push({ path: p, message: err?.message || String(err) });
      }
    }
  }

  // Remove DB rows for this jobsheet
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM scheduled_emails WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.scheduledRemoved = this.changes || 0; resolve(); }
    });
  });
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM email_log WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.emailsRemoved = this.changes || 0; resolve(); }
    });
  });
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM planner_actions WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.plannerRemoved = this.changes || 0; resolve(); }
    });
  });
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM documents WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.documentsRemoved = this.changes || 0; resolve(); }
    });
  });
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM jobsheet_template_overrides WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.overridesRemoved = this.changes || 0; resolve(); }
    });
  });
  await new Promise((resolve, reject) => {
    db.run(`DELETE FROM ahmen_jobsheets WHERE jobsheet_id = ?`, [jobsheetId], function (err) {
      if (err) reject(err); else { results.jobsheetsRemoved = this.changes || 0; resolve(); }
    });
  });

  return results;
}

function getAhmenVenues(options = {}) {
  const params = [];
  const conditions = [];
  if (options.businessId) {
    conditions.push('business_id = ?');
    params.push(options.businessId);
  }
  let query = 'SELECT * FROM ahmen_venues';
  if (conditions.length) {
    query += ` WHERE ${conditions.join(' AND ')}`;
  }
  query += ' ORDER BY LOWER(name) ASC';

  return new Promise((resolve, reject) => {
    db.all(query, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows.map(mapAhmenVenueRow));
    });
  });
}

function saveAhmenVenue(data) {
  return new Promise((resolve, reject) => {
    const businessId = Number(data?.business_id);
    if (!Number.isInteger(businessId)) {
      reject(new Error('Business id is required for venue'));
      return;
    }

    const name = (data?.name || '').trim();
    if (!name) {
      reject(new Error('Venue name is required'));
      return;
    }

    const insertValues = AHMEN_VENUE_FIELDS.map(field => sanitizeAhmenVenueValue(field, field === 'business_id' ? businessId : data?.[field]));
    const now = new Date().toISOString();

    if (data?.venue_id) {
      const venueId = Number(data.venue_id);
      if (!Number.isInteger(venueId)) {
        reject(new Error('Invalid venue id'));
        return;
      }

      const setClause = AHMEN_VENUE_FIELDS.map(field => `${field} = ?`).join(', ');
      db.run(
        `UPDATE ahmen_venues SET ${setClause}, updated_at = ? WHERE venue_id = ?`,
        [...insertValues, now, venueId],
        function (err) {
          if (err) reject(err);
          else resolve({ venue_id: venueId, changes: this.changes });
        }
      );
      return;
    }

    db.run(
      `INSERT INTO ahmen_venues (${AHMEN_VENUE_FIELDS.join(', ')}, created_at, updated_at) VALUES (${AHMEN_VENUE_FIELDS.map(() => '?').join(', ')}, ?, ?)`,
      [...insertValues, now, now],
      function (err) {
        if (err) reject(err);
        else resolve({ venue_id: this.lastID });
      }
    );
  });
}

function deleteAhmenVenue(venueId) {
  return new Promise((resolve, reject) => {
    const id = Number(venueId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid venue id'));
      return;
    }

    db.serialize(() => {
      db.run(
        'UPDATE ahmen_jobsheets SET venue_id = NULL WHERE venue_id = ?',
        [id],
        (updateErr) => {
          if (updateErr) {
            reject(updateErr);
            return;
          }
          db.run(
            'DELETE FROM ahmen_venues WHERE venue_id = ?',
            [id],
            function (err) {
              if (err) reject(err);
              else resolve(this.changes);
            }
          );
        }
      );
    });
  });
}

module.exports = {
  getInvoices: () => {
    return new Promise((resolve, reject) => {
      db.all(`
        SELECT invoices.invoice_number, clients.name as client, invoices.amount, invoices.due_date, invoices.status
        FROM invoices
        LEFT JOIN clients ON invoices.client_id = clients.client_id
        ORDER BY invoices.invoice_number
      `, [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  },

  getStatus: () => {
    return new Promise((resolve, reject) => {
      db.all(`
        SELECT invoices.status, COUNT(*) as count, SUM(invoices.amount) as total
        FROM invoices
        LEFT JOIN clients ON invoices.client_id = clients.client_id
        GROUP BY invoices.status
      `, [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  },

  getClients: () => {
    return new Promise((resolve, reject) => {
      db.all(`SELECT * FROM clients ORDER BY name`, [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  },

  getClientByName: (businessId, name) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const nm = (name || '').toString().trim();
      if (!Number.isInteger(id) || !nm) { resolve(null); return; }
      db.get(
        `SELECT * FROM clients WHERE business_id = ? AND lower(name) = lower(?) LIMIT 1`,
        [id, nm],
        (err, row) => {
          if (err) reject(err); else resolve(row || null);
        }
      );
    });
  },

  getClient: (clientId) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) { resolve(null); return; }
      db.get('SELECT * FROM clients WHERE client_id = ? LIMIT 1', [id], (err, row) => {
        if (err) reject(err); else resolve(row || null);
      });
    });
  },

  getClientDetails: (clientId) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) { resolve({ client: null, emails: [], phones: [], addresses: [] }); return; }
      const out = { client: null, emails: [], phones: [], addresses: [] };
      db.get('SELECT * FROM clients WHERE client_id = ? LIMIT 1', [id], (err, row) => {
        if (err) { reject(err); return; }
        out.client = row || null;
        db.all('SELECT * FROM client_emails WHERE client_id = ? ORDER BY is_primary DESC, id', [id], (e1, r1) => {
          if (e1) { reject(e1); return; }
          out.emails = Array.isArray(r1) ? r1 : [];
          db.all('SELECT * FROM client_phones WHERE client_id = ? ORDER BY is_primary DESC, id', [id], (e2, r2) => {
            if (e2) { reject(e2); return; }
            out.phones = Array.isArray(r2) ? r2 : [];
            db.all('SELECT * FROM client_addresses WHERE client_id = ? ORDER BY is_primary DESC, id', [id], (e3, r3) => {
              if (e3) { reject(e3); return; }
              out.addresses = Array.isArray(r3) ? r3 : [];
              resolve(out);
            });
          });
        });
      });
    });
  },

  markPaid: (invoiceNumber) => {
    return new Promise((resolve, reject) => {
      db.run(
        `UPDATE invoices SET status='Paid' WHERE invoice_number=?`,
        [invoiceNumber],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  resetStatus: (invoiceNumber) => {
    return new Promise((resolve, reject) => {
      db.run(
        `UPDATE invoices SET status='Issued' WHERE invoice_number=?`,
        [invoiceNumber],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  deleteInvoice: (invoiceNumber) => {
    return new Promise((resolve, reject) => {
      db.run(
        `DELETE FROM invoices WHERE invoice_number = ?`,
        [invoiceNumber],
        function (err) {
          if (err) reject(err);
          else resolve(this.changes);
        }
      );
    });
  },

  addInvoice: (clientId, amount, dueDate) => {
    return new Promise((resolve, reject) => {
      db.run(
        `INSERT INTO invoices (business_id, client_id, invoice_number, date_issued, due_date, amount, status)
         VALUES (${DEFAULT_BUSINESS_ID}, ?, (SELECT IFNULL(MAX(invoice_number), 0) + 1 FROM invoices WHERE business_id=${DEFAULT_BUSINESS_ID}), date('now'), ?, ?, 'Issued')`,
        [clientId, dueDate, amount],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  },

  addClient: (clientData) => {
    return new Promise((resolve, reject) => {
      const name = (clientData?.name || "").trim();
      if (!name) {
        reject(new Error("Client name is required"));
        return;
      }

      const email = clientData?.email?.trim() || null;
      const phone = clientData?.phone?.trim() || null;
      const contact = clientData?.contact?.trim() || null;
      const address = clientData?.address?.trim() || null;
      const address1 = clientData?.address1?.trim() || null;
      const address2 = clientData?.address2?.trim() || null;
      const town = clientData?.town?.trim() || null;
      const postcode = clientData?.postcode?.trim() || null;
      const businessId = clientData?.business_id ?? null;

      db.run(
        `INSERT INTO clients (business_id, name, email, phone, address, contact, address1, address2, town, postcode)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [businessId, name, email, phone, address, contact, address1, address2, town, postcode],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  },

  updateClient: (clientId, clientData) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) {
        reject(new Error("Invalid client id"));
        return;
      }

      const name = (clientData?.name || "").trim();
      if (!name) {
        reject(new Error("Client name is required"));
        return;
      }

      const email = clientData?.email?.trim() || null;
      const phone = clientData?.phone?.trim() || null;
      const contact = clientData?.contact?.trim() || null;
      const address = clientData?.address?.trim() || null;
      const address1 = clientData?.address1?.trim() || null;
      const address2 = clientData?.address2?.trim() || null;
      const town = clientData?.town?.trim() || null;
      const postcode = clientData?.postcode?.trim() || null;

      db.run(
        `UPDATE clients
         SET business_id=?, name=?, email=?, phone=?, address=?, contact=?, address1=?, address2=?, town=?, postcode=?
         WHERE client_id=?`,
        [clientData?.business_id ?? null, name, email, phone, address, contact, address1, address2, town, postcode, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  deleteClient: (clientId) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) { reject(new Error('Invalid client id')); return; }
      db.run('DELETE FROM clients WHERE client_id = ?', [id], function (err) {
        if (err) reject(err); else resolve(this.changes);
      });
    });
  },

  saveClientDetails: (clientId, details = {}) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) { reject(new Error('Invalid client id')); return; }

      const name = (details?.name || '').toString().trim();
      const emails = Array.isArray(details?.emails) ? details.emails.filter(e => e && e.email) : [];
      const phones = Array.isArray(details?.phones) ? details.phones.filter(p => p && p.phone) : [];
      const addresses = Array.isArray(details?.addresses) ? details.addresses.filter(a => a) : [];

      const normalizePrimary = (arr, key) => {
        let used = false;
        return arr.map(item => {
          const isP = item && (item.is_primary === 1 || item.is_primary === true || item.is_primary === '1');
          let flag = 0;
          if (isP && !used) { flag = 1; used = true; }
          return { ...item, [key]: (item && item[key]) || null, is_primary: flag };
        }).map((item, idx, list) => {
          if (!used && idx === 0) return { ...item, is_primary: 1 };
          return item;
        });
      };
      const emailsN = normalizePrimary(emails, 'email');
      const phonesN = normalizePrimary(phones, 'phone');
      const addressesN = (() => {
        let used = false;
        const list = addresses.map(a => ({
          label: (a?.label || '').toString().trim() || null,
          address1: (a?.address1 || '').toString().trim() || null,
          address2: (a?.address2 || '').toString().trim() || null,
          town: (a?.town || '').toString().trim() || null,
          postcode: (a?.postcode || '').toString().trim() || null,
          country: (a?.country || '').toString().trim() || null,
          is_primary: (a?.is_primary === 1 || a?.is_primary === true || a?.is_primary === '1') ? 1 : 0
        }));
        const any = list.some(it => it.is_primary === 1);
        if (!any && list.length) list[0].is_primary = 1;
        return list;
      })();

      // Derive base columns from primaries
      const primaryEmail = emailsN.find(e => e.is_primary === 1)?.email || null;
      const primaryPhone = phonesN.find(p => p.is_primary === 1)?.phone || null;
      const primaryAddr = addressesN.find(a => a.is_primary === 1) || {};
      const addressLine = [primaryAddr.address1, primaryAddr.address2, primaryAddr.town, primaryAddr.postcode, primaryAddr.country].filter(Boolean).join(', ');

      db.serialize(() => {
        // Update base client record
        db.run(
          `UPDATE clients SET name = COALESCE(?, name), email = ?, phone = ?, address = ?, address1 = ?, address2 = ?, town = ?, postcode = ? WHERE client_id = ?`,
          [name || null, primaryEmail, primaryPhone, addressLine || null, primaryAddr.address1 || null, primaryAddr.address2 || null, primaryAddr.town || null, primaryAddr.postcode || null, id],
          (err) => {
            if (err) { reject(err); return; }
            // Replace emails
            db.run('DELETE FROM client_emails WHERE client_id = ?', [id], (e1) => {
              if (e1) { reject(e1); return; }
              const insertEmail = db.prepare('INSERT INTO client_emails (client_id, label, email, is_primary, created_at, updated_at) VALUES (?, ?, ?, ?, datetime(\'now\'), datetime(\'now\'))');
              emailsN.forEach(item => insertEmail.run(id, item.label || null, item.email, item.is_primary ? 1 : 0));
              insertEmail.finalize((e1f) => {
                if (e1f) { reject(e1f); return; }
                // Replace phones
                db.run('DELETE FROM client_phones WHERE client_id = ?', [id], (e2) => {
                  if (e2) { reject(e2); return; }
                  const insertPhone = db.prepare('INSERT INTO client_phones (client_id, label, phone, is_primary, created_at, updated_at) VALUES (?, ?, ?, ?, datetime(\'now\'), datetime(\'now\'))');
                  phonesN.forEach(item => insertPhone.run(id, item.label || null, item.phone, item.is_primary ? 1 : 0));
                  insertPhone.finalize((e2f) => {
                    if (e2f) { reject(e2f); return; }
                    // Replace addresses
                    db.run('DELETE FROM client_addresses WHERE client_id = ?', [id], (e3) => {
                      if (e3) { reject(e3); return; }
                      const insertAddr = db.prepare('INSERT INTO client_addresses (client_id, label, address1, address2, town, postcode, country, is_primary, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime(\'now\'), datetime(\'now\'))');
                      addressesN.forEach(a => insertAddr.run(id, a.label || null, a.address1, a.address2, a.town, a.postcode, a.country, a.is_primary ? 1 : 0));
                      insertAddr.finalize((e3f) => {
                        if (e3f) { reject(e3f); return; }
                        resolve({ ok: true });
                      });
                    });
                  });
                });
              });
            });
          }
        );
      });
    });
  },

  getEvents: (options = {}) => {
    const params = [];
    const where = [];
    if (options.clientId) {
      where.push('events.client_id = ?');
      params.push(options.clientId);
    }
    if (options.businessId) {
      where.push('events.business_id = ?');
      params.push(options.businessId);
    }

    const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';

    return new Promise((resolve, reject) => {
      db.all(
        `SELECT events.*, clients.name AS client_name, business_settings.business_name
         FROM events
         LEFT JOIN clients ON events.client_id = clients.client_id
         LEFT JOIN business_settings ON events.business_id = business_settings.id
         ${whereClause}
         ORDER BY events.event_date DESC NULLS LAST, events.event_id DESC`,
        params,
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  },

  addEvent: (eventData) => {
    return new Promise((resolve, reject) => {
      const clientId = Number(eventData?.client_id);
      if (!Number.isInteger(clientId)) {
        reject(new Error('Client id is required for an event'));
        return;
      }

      const eventName = (eventData?.event_name || '').trim();
      if (!eventName) {
        reject(new Error('Event name is required'));
        return;
      }

      const eventDate = eventData?.event_date || null;
      const venueName = eventData?.venue_name || null;
      const venueAddress1 = eventData?.venue_address1 || null;
      const venueAddress2 = eventData?.venue_address2 || null;
      const town = eventData?.town || null;
      const postcode = eventData?.postcode || null;
      const notes = eventData?.notes || null;
      const businessId = eventData?.business_id || null;

      db.run(
        `INSERT INTO events (client_id, business_id, event_name, event_date, venue_name, venue_address1, venue_address2, town, postcode, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ,
        [clientId, businessId, eventName, eventDate, venueName, venueAddress1, venueAddress2, town, postcode, notes],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  },

  updateEvent: (eventId, eventData) => {
    return new Promise((resolve, reject) => {
      const id = Number(eventId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid event id'));
        return;
      }

      const eventName = (eventData?.event_name || '').trim();
      if (!eventName) {
        reject(new Error('Event name is required'));
        return;
      }

      const eventDate = eventData?.event_date || null;
      const venueName = eventData?.venue_name || null;
      const venueAddress1 = eventData?.venue_address1 || null;
      const venueAddress2 = eventData?.venue_address2 || null;
      const town = eventData?.town || null;
      const postcode = eventData?.postcode || null;
      const notes = eventData?.notes || null;
      const businessId = eventData?.business_id || null;

      db.run(
        `UPDATE events
         SET client_id = ?, business_id = ?, event_name = ?, event_date = ?, venue_name = ?, venue_address1 = ?, venue_address2 = ?, town = ?, postcode = ?, notes = ?
         WHERE event_id = ?`,
        [eventData?.client_id, businessId, eventName, eventDate, venueName, venueAddress1, venueAddress2, town, postcode, notes, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  getDocuments: (options = {}) => {
    const params = [];
    const where = [];

    if (options.docType) {
      where.push('documents.doc_type = ?');
      params.push(options.docType);
    }
    if (options.businessId) {
      where.push('documents.business_id = ?');
      params.push(options.businessId);
    }
    if (options.eventId) {
      where.push('documents.event_id = ?');
      params.push(options.eventId);
    }
    if (options.jobsheetId) {
      where.push('documents.jobsheet_id = ?');
      params.push(options.jobsheetId);
    }

    const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';

    return new Promise((resolve, reject) => {
      db.all(
        `SELECT
           documents.*,
           def.label AS definition_label,
           def.invoice_variant AS definition_invoice_variant,
           def.doc_type AS definition_doc_type,
           COALESCE(documents.client_name, clients.name) AS display_client_name,
           COALESCE(documents.event_name, events.event_name) AS display_event_name,
           COALESCE(documents.event_date, events.event_date) AS display_event_date,
           events.event_name AS joined_event_name,
           events.event_date AS joined_event_date,
           clients.name AS joined_client_name,
           business_settings.business_name
         FROM documents
         LEFT JOIN document_definitions def ON def.business_id = documents.business_id AND def.key = documents.definition_key
         LEFT JOIN events ON documents.event_id = events.event_id
         LEFT JOIN clients ON events.client_id = clients.client_id
         LEFT JOIN business_settings ON documents.business_id = business_settings.id
         ${whereClause}
         ORDER BY documents.created_at DESC, documents.document_id DESC`,
        params,
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  },

  getDocumentDefinitions: (businessId, options) => getDocumentDefinitionsForBusiness(businessId, options),
  getDocumentDefinition: (businessId, identifier) => getDocumentDefinitionRecord(businessId, identifier),
  saveDocumentDefinition: (businessId, definition) => saveDocumentDefinition(businessId, definition),
  deleteDocumentDefinition: (businessId, identifier) => deleteDocumentDefinition(businessId, identifier),
  reorderDocumentDefinitions: (businessId, orderedKeys) => reorderDocumentDefinitions(businessId, orderedKeys),
  getJobsheetTemplateOverrides: (jobsheetId) => getJobsheetTemplateOverrides(jobsheetId),
  setJobsheetTemplateOverride: (jobsheetId, definitionKey, templatePath) => setJobsheetTemplateOverride(jobsheetId, definitionKey, templatePath),
  clearJobsheetTemplateOverride: (jobsheetId, definitionKey) => clearJobsheetTemplateOverride(jobsheetId, definitionKey),

  addDocument: (documentData) => {
    return new Promise((resolve, reject) => {
      const docType = (documentData?.doc_type || '').toLowerCase();
      if (!docType) {
        reject(new Error('Document type is required'));
        return;
      }

      const businessId = documentData?.business_id || null;
      const eventId = documentData?.event_id || null;
      const rawJobsheetId = documentData?.jobsheet_id;
      const jobsheetId = rawJobsheetId != null ? Number(rawJobsheetId) : null;
      const normalizedJobsheetId = Number.isInteger(jobsheetId) ? jobsheetId : null;
      const status = documentData?.status || 'draft';
      const totalAmount = documentData?.total_amount || 0;
      const balanceDue = documentData?.balance_due ?? totalAmount;
      const dueDate = documentData?.due_date || null;
      const filePath = documentData?.file_path || null;
      const clientName = documentData?.client_name || null;
      const eventName = documentData?.event_name || null;
      const eventDate = documentData?.event_date || null;
      const documentDate = documentData?.document_date || null;
      const definitionKey = documentData?.definition_key || documentData?.document_definition_key || null;
      const invoiceVariant = documentData?.invoice_variant || null;

      const requestedNumber = documentData?.number ? Number(documentData.number) : null;
      const counterColumn = getCounterColumn(docType);

      const finalizeInsert = (resolvedNumber) => {
        db.run(
          `INSERT INTO documents (
             event_id,
             jobsheet_id,
             business_id,
             doc_type,
             number,
             status,
             total_amount,
             balance_due,
             due_date,
             file_path,
             client_name,
             event_name,
             event_date,
             document_date,
             definition_key,
             invoice_variant
           ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ,
          [
            eventId,
            normalizedJobsheetId,
            businessId,
            docType,
            resolvedNumber,
            status,
            totalAmount,
            balanceDue,
            dueDate,
            filePath,
            clientName,
            eventName,
            eventDate,
            documentDate,
            definitionKey,
            invoiceVariant
          ],
          function (err) {
            if (err) {
              reject(err);
            } else {
              resolve({ id: this.lastID, number: resolvedNumber });
            }
          }
        );
      };

      if (requestedNumber != null) {
        if (!businessId) {
          reject(new Error('Business is required when specifying a document number'));
          return;
        }

        db.get(
          `SELECT document_id FROM documents WHERE business_id = ? AND doc_type = ? AND number = ? LIMIT 1`,
          [businessId, docType, requestedNumber],
          (dupErr, existing) => {
            if (dupErr) {
              reject(dupErr);
              return;
            }
            if (existing) {
              reject(new Error('Document number already exists for this business and document type'));
              return;
            }
            finalizeInsert(requestedNumber);
          }
        );
        return;
      }

      if (!counterColumn || !businessId) {
        finalizeInsert(null);
        return;
      }

      db.get(
        `SELECT ${counterColumn} AS counter FROM business_settings WHERE id = ?`,
        [businessId],
        (err, row) => {
          if (err) {
            reject(err);
            return;
          }

          const nextNumber = (row?.counter || 0) + 1;

          db.run(
            `UPDATE business_settings SET ${counterColumn} = ? WHERE id = ?`,
            [nextNumber, businessId],
            (updateErr) => {
              if (updateErr) {
                reject(updateErr);
                return;
              }
              finalizeInsert(nextNumber);
            }
          );
        }
      );
    });
  },

  // Replace all line items for a document (simple sync)
  saveDocumentItems: (documentId, items = []) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) { reject(new Error('Invalid document id')); return; }
      const list = Array.isArray(items) ? items : [];
      db.serialize(() => {
        db.run('DELETE FROM document_items WHERE document_id = ?', [id], (delErr) => {
          if (delErr) { reject(delErr); return; }
          if (!list.length) { resolve({ inserted: 0 }); return; }
          const stmt = db.prepare(`INSERT INTO document_items (document_id, item_type, description, quantity, unit, rate, amount, sort_order, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))`);
          let count = 0;
          list.forEach((raw, idx) => {
            const type = (raw?.item_type || raw?.type || '').toString().trim().toLowerCase() || null;
            const desc = (raw?.description || '').toString().trim() || null;
            const qty = Number(raw?.quantity);
            const unit = (raw?.unit || '').toString().trim() || null;
            const rate = Number(raw?.rate);
            const amount = Number.isFinite(Number(raw?.amount)) ? Number(raw?.amount) : (Number.isFinite(qty) && Number.isFinite(rate) ? qty * rate : null);
            stmt.run(id, type, desc, Number.isFinite(qty) ? qty : null, unit, Number.isFinite(rate) ? rate : null, Number.isFinite(amount) ? amount : null, Number.isInteger(raw?.sort_order) ? Number(raw.sort_order) : idx);
            count += 1;
          });
          stmt.finalize((finErr) => finErr ? reject(finErr) : resolve({ inserted: count }));
        });
      });
    });
  },

  getDocumentItems: (documentId) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) { resolve([]); return; }
      db.all('SELECT * FROM document_items WHERE document_id = ? ORDER BY sort_order, item_id', [id], (err, rows) => {
        if (err) reject(err); else resolve(Array.isArray(rows) ? rows : []);
      });
    });
  },

  updateDocumentStatus: (documentId, data = {}) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid document id'));
        return;
      }

      const updates = [];
      const params = [];

      if (data.status !== undefined) {
        updates.push('status = ?');
        params.push(data.status);
      }
      if (data.total_amount !== undefined) {
        updates.push('total_amount = ?');
        params.push(data.total_amount);
      }
      if (data.balance_due !== undefined) {
        updates.push('balance_due = ?');
        params.push(data.balance_due);
      }
      if (data.due_date !== undefined) {
        updates.push('due_date = ?');
        params.push(data.due_date);
      }
      if (data.reminder_date !== undefined) {
        updates.push('reminder_date = ?');
        params.push(data.reminder_date);
      }
      if (data.paid_at !== undefined) {
        updates.push('paid_at = ?');
        params.push(data.paid_at);
      }
      if (data.file_path !== undefined) {
        updates.push('file_path = ?');
        params.push(data.file_path);
      }
      if (data.definition_key !== undefined) {
        updates.push('definition_key = ?');
        params.push(data.definition_key);
      }
      if (data.doc_type !== undefined) {
        updates.push('doc_type = ?');
        params.push(data.doc_type);
      }
      if (data.jobsheet_id !== undefined) {
        const jid = data.jobsheet_id == null ? null : Number(data.jobsheet_id);
        updates.push('jobsheet_id = ?');
        params.push(Number.isInteger(jid) ? jid : null);
      }
      if (data.invoice_variant !== undefined) {
        updates.push('invoice_variant = ?');
        params.push(data.invoice_variant);
      }
      // Allow editing identity fields visible in UI
      if (data.client_name !== undefined) {
        updates.push('client_name = ?');
        params.push(data.client_name);
      }
      if (data.event_name !== undefined) {
        updates.push('event_name = ?');
        params.push(data.event_name);
      }
      if (data.event_date !== undefined) {
        updates.push('event_date = ?');
        params.push(data.event_date);
      }

      if (!updates.length) {
        resolve();
        return;
      }

      params.push(id);

      db.run(
        `UPDATE documents SET ${updates.join(', ')}, updated_at = datetime('now') WHERE document_id = ?`,
        params,
        function (err) {
          if (err) {
            reject(err);
          } else {
            updateDocumentTimestamp(id);
            resolve();
          }
        }
      );
    });
  },

  getMusiciansForEvent: (eventId) => {
    return new Promise((resolve, reject) => {
      db.all(
        `SELECT * FROM event_musicians WHERE event_id = ? ORDER BY musician_id`,
        [eventId],
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  },

  addMusicianToEvent: (eventId, musicianData) => {
    return new Promise((resolve, reject) => {
      const name = (musicianData?.name || '').trim();
      if (!name) {
        reject(new Error('Musician name is required'));
        return;
      }

      const role = musicianData?.role || null;
      const fee = Number(musicianData?.fee || 0);
      const paidStatus = musicianData?.paid_status || 'unpaid';

      db.run(
        `INSERT INTO event_musicians (event_id, name, role, fee, paid_status)
         VALUES (?, ?, ?, ?, ?)`,
        [eventId, name, role, fee, paidStatus],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  },

  updateMusicianPayment: (musicianId, data = {}) => {
    return new Promise((resolve, reject) => {
      const id = Number(musicianId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid musician id'));
        return;
      }

      const updates = [];
      const params = [];

      if (data.paid_status !== undefined) {
        updates.push('paid_status = ?');
        params.push(data.paid_status);
      }
      if (data.fee !== undefined) {
        updates.push('fee = ?');
        params.push(data.fee);
      }

      if (!updates.length) {
        resolve();
        return;
      }

      params.push(id);

      db.run(
        `UPDATE event_musicians SET ${updates.join(', ')} WHERE musician_id = ?`,
        params,
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  getTimekeeperSessions: (options = {}) => {
    const params = [];
    const where = [];

    if (options.clientId) {
      where.push('client_id = ?');
      params.push(options.clientId);
    }
    if (options.eventId) {
      where.push('event_id = ?');
      params.push(options.eventId);
    }
    if (options.exported !== undefined) {
      where.push('exported = ?');
      params.push(options.exported ? 1 : 0);
    }

    const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';

    return new Promise((resolve, reject) => {
      db.all(
        `SELECT * FROM timekeeper_sessions ${whereClause} ORDER BY session_date DESC, session_id DESC`,
        params,
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  },

  importTimekeeperSession: (sessionData) => {
    return new Promise((resolve, reject) => {
      const description = (sessionData?.description || '').trim();
      if (!description) {
        reject(new Error('Session description is required'));
        return;
      }

      const clientId = sessionData?.client_id || null;
      const eventId = sessionData?.event_id || null;
      const sessionDate = sessionData?.session_date || null;
      const durationMinutes = sessionData?.duration_minutes || null;
      const rate = sessionData?.rate || null;
      const amount = sessionData?.amount || null;
      const exported = sessionData?.exported ? 1 : 0;

      db.run(
        `INSERT INTO timekeeper_sessions (client_id, event_id, description, session_date, duration_minutes, rate, amount, exported)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        ,
        [clientId, eventId, description, sessionDate, durationMinutes, rate, amount, exported],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  },

  markSessionExported: (sessionId, exported = true) => {
    return new Promise((resolve, reject) => {
      const id = Number(sessionId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid session id'));
        return;
      }

      db.run(
        `UPDATE timekeeper_sessions SET exported = ? WHERE session_id = ?`,
        [exported ? 1 : 0, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  businessSettings: () => {
    return new Promise((resolve, reject) => {
      db.all(
        `SELECT * FROM business_settings ORDER BY id`,
        [],
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  },

  updateBusinessSettings: (businessId, updates) => updateBusinessSettingsRecord(businessId, updates),

  getClientById: (clientId) => {
    return new Promise((resolve, reject) => {
      db.get(
        `SELECT * FROM clients WHERE client_id = ?`,
        [clientId],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });
  },

  getEventById: (eventId) => {
    return new Promise((resolve, reject) => {
      db.get(
        `SELECT * FROM events WHERE event_id = ?`,
        [eventId],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });
  },

  getBusinessById: (businessId) => fetchBusinessRecord(businessId),

  getDocumentById: (documentId) => {
    return new Promise((resolve, reject) => {
      db.get(
        `SELECT * FROM documents WHERE document_id = ?`,
        [documentId],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });
  },

  getDocumentByFilePath: (businessId, filePath) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      if (!Number.isInteger(id) || !filePath) {
        resolve(null);
        return;
      }

      db.get(
        `SELECT * FROM documents WHERE business_id = ? AND file_path = ? LIMIT 1`,
        [id, filePath],
        (err, row) => {
          if (err) reject(err);
          else resolve(row || null);
        }
      );
    });
  },
  getMaxInvoiceNumber: (businessId) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      if (!Number.isInteger(id)) {
        resolve(null);
        return;
      }
      db.get(
        `SELECT MAX(number) AS maxnum FROM documents WHERE business_id = ? AND lower(doc_type) = 'invoice'`,
        [id],
        (err, row) => {
          if (err) reject(err);
          else resolve(row?.maxnum != null ? Number(row.maxnum) : null);
        }
      );
    });
  },
  // Generic max number getter for a doc type (invoice/quote)
  getMaxNumberForDocType: (businessId, docType) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const type = (docType || '').toString().toLowerCase();
      if (!Number.isInteger(id) || !type) { resolve(null); return; }
      db.get(
        `SELECT MAX(number) AS maxnum FROM documents WHERE business_id = ? AND lower(doc_type) = ?`,
        [id, type],
        (err, row) => {
          if (err) reject(err);
          else resolve(row?.maxnum != null ? Number(row.maxnum) : null);
        }
      );
    });
  },
  getDocumentsByNumber: (businessId, docType, number) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const type = (docType || '').toString().toLowerCase();
      const num = Number(number);
      if (!Number.isInteger(id) || !type || !Number.isInteger(num)) { resolve([]); return; }
      db.all(
        `SELECT * FROM documents WHERE business_id = ? AND lower(doc_type) = ? AND number = ?`,
        [id, type, num],
        (err, rows) => {
          if (err) reject(err);
          else resolve(Array.isArray(rows) ? rows : []);
        }
      );
    });
  },
  documentNumberExists: (businessId, docType, number) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const type = (docType || '').toString().toLowerCase();
      const num = Number(number);
      if (!Number.isInteger(id) || !type || !Number.isInteger(num)) { resolve(false); return; }
      db.all(
        `SELECT document_id, file_path FROM documents WHERE business_id = ? AND lower(doc_type) = ? AND number = ?`,
        [id, type, num],
        async (err, rows) => {
          if (err) { reject(err); return; }
          const list = Array.isArray(rows) ? rows : [];
          if (!list.length) { resolve(false); return; }
          try {
            // If any has an existing file, it's taken. Otherwise, purge orphans and report available.
            const exists = list.some(r => r && r.file_path && fs.existsSync(r.file_path));
            if (exists) { resolve(true); return; }
            // Clean up phantom rows (no existing file)
            await Promise.all(list.map(r => new Promise((res) => {
              db.run(`DELETE FROM documents WHERE document_id = ?`, [r.document_id], () => res());
            })));
            resolve(false);
          } catch (e) {
            resolve(false);
          }
        }
      );
    });
  },
  setLastInvoiceNumber: (businessId, nextVal) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const val = Number(nextVal);
      if (!Number.isInteger(id)) { reject(new Error('Invalid business id')); return; }
      if (!Number.isInteger(val) || val < 0) { reject(new Error('Invalid invoice number')); return; }
      db.run(
        `UPDATE business_settings SET last_invoice_number = ? WHERE id = ?`,
        [val, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },
  // Generic setter for last number counter by doc type
  setLastNumberForDocType: (businessId, docType, nextVal) => {
    return new Promise((resolve, reject) => {
      const id = Number(businessId);
      const type = (docType || '').toString().toLowerCase();
      const val = Number(nextVal);
      if (!Number.isInteger(id)) { reject(new Error('Invalid business id')); return; }
      if (!Number.isInteger(val) || val < 0) { reject(new Error('Invalid number')); return; }
      const column = getCounterColumn(type);
      if (!column) { reject(new Error('Unsupported document type')); return; }
      db.run(
        `UPDATE business_settings SET ${column} = ? WHERE id = ?`,
        [val, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  // Promote a pdf_export (or any doc) to a real invoice row with assigned number.
  // Copies metadata and links the existing PDF path. If an invoice already exists for the same file_path,
  // returns that invoice without creating a duplicate.
  promotePdfToInvoice: (documentId, options = {}) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) { reject(new Error('Invalid document id')); return; }

      db.get(`SELECT * FROM documents WHERE document_id = ?`, [id], async (err, row) => {
        if (err) { reject(err); return; }
        if (!row) { reject(new Error('Document not found')); return; }

        try {
          const doc = row;
          const businessId = Number(doc.business_id);
          if (!Number.isInteger(businessId)) { reject(new Error('Invalid business id for document')); return; }
          const filePath = doc.file_path || null;
          if (!filePath) { reject(new Error('Document file path is required to promote')); return; }

          // If an invoice record already exists for this path, return it
          const existingInvoice = await new Promise((res, rej) => {
            db.get(
              `SELECT * FROM documents WHERE business_id = ? AND file_path = ? AND lower(doc_type) = 'invoice' LIMIT 1`,
              [businessId, filePath],
              (e, r) => e ? rej(e) : res(r || null)
            );
          });
          if (existingInvoice) { resolve({ id: existingInvoice.document_id, number: existingInvoice.number }); return; }

          const jobsheetId = doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null;
          let js = null;
          if (Number.isInteger(jobsheetId)) {
            try {
              js = await new Promise((res, rej) => {
                db.get(`SELECT * FROM ahmen_jobsheets WHERE jobsheet_id = ?`, [jobsheetId], (e, r) => e ? rej(e) : res(r || null));
              });
            } catch (_) { js = null; }
          }

          const lower = (v) => (v == null ? '' : String(v).toLowerCase());
          const inferVariant = () => {
            const fromDoc = lower(doc.definition_invoice_variant || doc.invoice_variant);
            if (fromDoc === 'deposit' || fromDoc === 'balance') return fromDoc;
            const pathName = lower(doc.file_path || '');
            const labelName = lower(doc.definition_label || doc.label || '');
            const hay = `${pathName} ${labelName}`;
            if (hay.includes('deposit')) return 'deposit';
            if (hay.includes('balance')) return 'balance';
            return '';
          };
          const variant = inferVariant();

          // Derive amounts/dates
          let totalAmount = doc.total_amount != null ? Number(doc.total_amount) : null;
          let balanceDue = doc.balance_due != null ? Number(doc.balance_due) : null;
          let dueDate = doc.due_date || null;
          let reminderDate = doc.reminder_date || null;

          if (js) {
            if (variant === 'deposit') {
              totalAmount = Number.isFinite(Number(js.deposit_amount)) ? Number(js.deposit_amount) : totalAmount;
              // No reminder for deposit
              reminderDate = null;
              // Use event_date as due if no explicit due date
              if (!dueDate) dueDate = js.event_date || null;
            } else if (variant === 'balance') {
              totalAmount = Number.isFinite(Number(js.balance_amount)) ? Number(js.balance_amount) : totalAmount;
              balanceDue = totalAmount;
              if (!dueDate) dueDate = js.balance_due_date || null;
              if (reminderDate == null && js.balance_reminder_date != null) reminderDate = js.balance_reminder_date;
            }
          }

          // Compose insert via addDocument to get an assigned number
          const requestedNumber = options && options.number != null ? Number(options.number) : null;
          const payload = {
            business_id: businessId,
            jobsheet_id: jobsheetId,
            doc_type: 'invoice',
            number: requestedNumber != null && Number.isInteger(requestedNumber) ? requestedNumber : undefined,
            status: 'issued',
            total_amount: totalAmount,
            balance_due: balanceDue != null ? balanceDue : totalAmount,
            due_date: dueDate,
            file_path: filePath,
            client_name: doc.client_name || null,
            event_name: doc.event_name || null,
            event_date: doc.event_date || null,
            document_date: doc.document_date || new Date().toISOString(),
            definition_key: doc.definition_key || null,
            invoice_variant: variant || null
          };

          try {
            const inserted = await module.exports.addDocument(payload);
            if (reminderDate != null) {
              try { await module.exports.updateDocumentStatus(inserted.id, { reminder_date: reminderDate }); } catch(_){}
            }
            resolve({ id: inserted.id, number: inserted.number });
          } catch (e) {
            // Fallback: if requested number collides, insert without a number and then set it explicitly
            try {
              const fallback = { ...payload };
              delete fallback.number;
              const ins2 = await module.exports.addDocument(fallback);
              // Assign requested number, allowing duplicates per relaxed rule
              if (requestedNumber != null && Number.isInteger(requestedNumber)) {
                try { await module.exports.setDocumentNumber(ins2.id, requestedNumber); } catch (_) {}
              }
              if (reminderDate != null) {
                try { await module.exports.updateDocumentStatus(ins2.id, { reminder_date: reminderDate }); } catch(_){}
              }
              resolve({ id: ins2.id, number: requestedNumber != null ? requestedNumber : ins2.number });
            } catch (ex2) {
              reject(e);
            }
          }
        } catch (ex) {
          reject(ex);
        }
      });
    });
  },

  // Set/override a specific invoice's number with validation and counter sync
  setDocumentNumber: (documentId, newNumber) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      const num = Number(newNumber);
      if (!Number.isInteger(id)) { reject(new Error('Invalid document id')); return; }
      if (!Number.isInteger(num) || num < 0) { reject(new Error('Invalid invoice number')); return; }

      db.get(
        `SELECT document_id, business_id, doc_type FROM documents WHERE document_id = ?`,
        [id],
        (err, row) => {
          if (err) { reject(err); return; }
          if (!row) { reject(new Error('Document not found')); return; }
          const docType = String(row.doc_type || '').toLowerCase();
          if (docType !== 'invoice') { reject(new Error('Only invoice documents can be renumbered')); return; }
          const businessId = row.business_id;
          if (!Number.isInteger(businessId)) { reject(new Error('Invalid business id for document')); return; }

          // Allow duplicate numbers for historical cases; update directly and keep counter in sync
          db.run(
            `UPDATE documents SET number = ?, updated_at = datetime('now') WHERE document_id = ?`,
            [num, id],
            function (updateErr) {
              if (updateErr) { reject(updateErr); return; }
              // Sync the last_invoice_number to the current max for safety
              db.get(
                `SELECT MAX(number) AS maxnum FROM documents WHERE business_id = ? AND lower(doc_type) = 'invoice'`,
                [businessId],
                (maxErr, maxRow) => {
                  if (maxErr) { /* ignore sync error */ resolve(); return; }
                  const maxNum = maxRow?.maxnum != null ? Number(maxRow.maxnum) : 0;
                  db.run(
                    `UPDATE business_settings SET last_invoice_number = ? WHERE id = ?`,
                    [maxNum, businessId],
                    function (_syncErr) { resolve(); }
                  );
                }
              );
            }
          );
        }
      );
    });
  },

  setDocumentLock: (documentId, locked) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid document id'));
        return;
      }
      const value = locked ? 1 : 0;
      db.run(
        `UPDATE documents SET is_locked = ?, updated_at = datetime('now') WHERE document_id = ?`,
        [value, id],
        function (err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });
  },

  deleteDocument: (documentId) => {
    return new Promise((resolve, reject) => {
      const id = Number(documentId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid document id'));
        return;
      }

      db.run(
        `DELETE FROM documents WHERE document_id = ?`,
        [id],
        function (err) {
          if (err) reject(err);
          else resolve(this.changes);
        }
      );
    });
  },
  // Expose utility to update a single document's file_path (used for dedup cleanup)
  setDocumentFilePath: (documentId, filePath) => setDocumentFilePath(documentId, filePath),
  clearDocumentPath: (businessId, absolutePath) => clearDocumentPath(businessId, absolutePath),

  deleteMusician: (musicianId) => {
    return new Promise((resolve, reject) => {
      const id = Number(musicianId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid musician id'));
        return;
      }

      db.run(
        `DELETE FROM event_musicians WHERE musician_id = ?`,
        [id],
        function (err) {
          if (err) reject(err);
          else resolve(this.changes);
        }
      );
    });
  },

  deleteTimekeeperSession: (sessionId, options = {}) => {
    return new Promise((resolve, reject) => {
      const id = Number(sessionId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid session id'));
        return;
      }

      db.get(
        `SELECT exported FROM timekeeper_sessions WHERE session_id = ?`,
        [id],
        (err, row) => {
          if (err) {
            reject(err);
            return;
          }
          if (!row) {
            resolve(0);
            return;
          }
          if (row.exported && !options.force) {
            reject(new Error('Session has already been exported. Pass force=true to delete.'));
            return;
          }

          db.run(
            `DELETE FROM timekeeper_sessions WHERE session_id = ?`,
            [id],
            function (deleteErr) {
              if (deleteErr) reject(deleteErr);
              else resolve(this.changes);
            }
          );
        }
      );
    });
  },

  deleteEvent: (eventId) => {
    return new Promise((resolve, reject) => {
      const id = Number(eventId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid event id'));
        return;
      }

      db.get(`SELECT event_id FROM events WHERE event_id = ?`, [id], (findErr, row) => {
        if (findErr) {
          reject(findErr);
          return;
        }
        if (!row) {
          resolve(0);
          return;
        }

        db.get(`SELECT COUNT(*) AS count FROM documents WHERE event_id = ?`, [id], (docErr, docRow) => {
          if (docErr) {
            reject(docErr);
            return;
          }
          if (docRow?.count > 0) {
            reject(new Error('Cannot delete event while documents reference it. Delete documents first.'));
            return;
          }

          db.get(`SELECT COUNT(*) AS count FROM timekeeper_sessions WHERE event_id = ?`, [id], (sessionErr, sessionRow) => {
            if (sessionErr) {
              reject(sessionErr);
              return;
            }
            if (sessionRow?.count > 0) {
              reject(new Error('Cannot delete event while Timekeeper sessions reference it. Delete sessions first.'));
              return;
            }

            db.serialize(() => {
              db.run('BEGIN TRANSACTION');
              db.run(
                `DELETE FROM event_musicians WHERE event_id = ?`,
                [id],
                function (musErr) {
                  if (musErr) {
                    db.run('ROLLBACK');
                    reject(musErr);
                    return;
                  }

                  db.run(
                    `DELETE FROM events WHERE event_id = ?`,
                    [id],
                    function (eventErr) {
                      if (eventErr) {
                        db.run('ROLLBACK');
                        reject(eventErr);
                        return;
                      }
                      const deletedEvents = this.changes;

                      db.run('COMMIT', commitErr => {
                        if (commitErr) {
                          reject(commitErr);
                          return;
                        }
                        resolve(deletedEvents);
                      });
                    }
                  );
                }
              );
            });
          });
        });
      });
    });
  },

  deleteClient: (clientId) => {
    return new Promise((resolve, reject) => {
      const id = Number(clientId);
      if (!Number.isInteger(id)) {
        reject(new Error('Invalid client id'));
        return;
      }

      db.get(`SELECT COUNT(*) AS count FROM events WHERE client_id = ?`, [id], (err, row) => {
        if (err) {
          reject(err);
          return;
        }
        if (row?.count > 0) {
          reject(new Error('Cannot delete client while events exist. Delete associated events first.'));
          return;
        }

        db.run(
          `DELETE FROM clients WHERE client_id = ?`,
          [id],
          function (deleteErr) {
            if (deleteErr) reject(deleteErr);
            else resolve(this.changes);
          }
        );
      });
    });
  },

  getAhmenJobsheets,
  getAhmenJobsheet,
  addAhmenJobsheet,
  updateAhmenJobsheet,
  updateAhmenJobsheetStatus,
  setJobsheetArchived,
  deleteAhmenJobsheet,
  getAhmenVenues,
  saveAhmenVenue,
  deleteAhmenVenue,
  getMergeFields,
  getMergeFieldBindingsByTemplate,
  getMergeFieldValueSources,
  setMergeFieldValueSource,
  clearMergeFieldValueSource,
  saveMergeField,
  deleteMergeField,
  deleteDocumentsByDefinition,
  deleteDocumentsByPathPrefix,
  deleteDocumentByFilePath,
  relocateBusinessDocuments,
  listDocumentTree,
  deleteDocumentFolder,
  deleteDocumentByPath,
  deleteJobsheetCompletely,
emptyDocumentsTrash
};

// Email log helpers (exported after main object)
module.exports.logEmail = ({
  business_id = null,
  jobsheet_id = null,
  to,
  cc,
  bcc,
  subject,
  body,
  attachments = [],
  provider = 'graph',
  status = 'sent',
  message_id = null,
  sent_at = null
}) => {
  return new Promise((resolve, reject) => {
    const toAddr = (to || '').toString();
    if (!toAddr) { reject(new Error('to is required')); return; }
    const ccAddr = Array.isArray(cc) ? cc.join(', ') : (cc || '');
    const bccAddr = Array.isArray(bcc) ? bcc.join(', ') : (bcc || '');
    const atts = JSON.stringify(Array.isArray(attachments) ? attachments : (attachments ? [attachments] : []));
    const sentAt = toSqliteDateTime(sent_at);
    db.run(
      `INSERT INTO email_log (business_id, jobsheet_id, to_address, cc_address, bcc_address, subject, body, attachments, provider, status, message_id, sent_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, COALESCE(?, datetime('now')))`,
      [business_id, jobsheet_id, toAddr, ccAddr, bccAddr, subject || '', body || '', atts, provider, status, message_id || null, sentAt],
      function (err) {
        if (err) reject(err);
        else resolve(this.lastID);
      }
    );
  });
};

module.exports.listEmailLog = ({ business_id = null, jobsheet_id = null, limit = 100 } = {}) => {
  return new Promise((resolve, reject) => {
    const where = [];
    const params = [];
    if (business_id != null) { where.push('business_id = ?'); params.push(business_id); }
    if (jobsheet_id != null) { where.push('jobsheet_id = ?'); params.push(jobsheet_id); }
    const sql = `SELECT * FROM email_log ${where.length ? 'WHERE ' + where.join(' AND ') : ''} ORDER BY sent_at DESC, id DESC LIMIT ?`;
    params.push(Number(limit) || 100);
    db.all(sql, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows || []);
    });
  });
};

module.exports.deleteEmailLog = (id) => {
  return new Promise((resolve, reject) => {
    const numericId = Number(id);
    if (!Number.isInteger(numericId) || numericId <= 0) {
      reject(new Error('Invalid email log id'));
      return;
    }
    db.serialize(() => {
      db.run('DELETE FROM scheduled_emails WHERE email_log_id = ?', [numericId], () => {});
      db.run('DELETE FROM email_log WHERE id = ?', [numericId], function (err) {
        if (err) {
          reject(err);
        } else {
          resolve({ deleted: this.changes || 0 });
        }
      });
    });
  });
};

module.exports.updateEmailLogStatus = ({ id, status, sent_at = null }) => {
  return new Promise((resolve, reject) => {
    const numericId = Number(id);
    if (!Number.isInteger(numericId) || numericId <= 0) {
      reject(new Error('Invalid email log id'));
      return;
    }
    const sentAt = toSqliteDateTime(sent_at);
    db.run(
      `UPDATE email_log SET status = ?, sent_at = CASE WHEN ? IS NOT NULL THEN ? ELSE sent_at END WHERE id = ?`,
      [status || 'sent', sentAt, sentAt, numericId],
      function (err) {
        if (err) reject(err);
        else resolve({ updated: this.changes || 0 });
      }
    );
  });
};

module.exports.queueScheduledEmail = ({
  email_log_id = null,
  business_id = null,
  jobsheet_id = null,
  to,
  cc,
  bcc,
  subject,
  body,
  attachments = [],
  is_html = true,
  send_at
}) => {
  return new Promise((resolve, reject) => {
    const toAddr = (to || '').toString().trim();
    if (!toAddr) { reject(new Error('to is required')); return; }
    const sendAt = toSqliteDateTime(send_at);
    if (!sendAt) { reject(new Error('send_at is invalid')); return; }
    const ccAddr = Array.isArray(cc) ? cc.join(', ') : (cc || '');
    const bccAddr = Array.isArray(bcc) ? bcc.join(', ') : (bcc || '');
    const atts = JSON.stringify(Array.isArray(attachments) ? attachments : (attachments ? [attachments] : []));
    db.run(
      `INSERT INTO scheduled_emails (email_log_id, business_id, jobsheet_id, to_address, cc_address, bcc_address, subject, body, attachments, is_html, send_at, status, attempt_count, last_error, sent_at, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 0, NULL, NULL, datetime('now'), datetime('now'))`,
      [email_log_id, business_id, jobsheet_id, toAddr, ccAddr, bccAddr, subject || '', body || '', atts, is_html ? 1 : 0, sendAt],
      function (err) {
        if (err) reject(err);
        else resolve(this.lastID);
      }
    );
  });
};

module.exports.listDueScheduledEmails = ({ limit = 10 } = {}) => {
  return new Promise((resolve, reject) => {
    const cap = Number(limit) || 10;
    db.all(
      `SELECT * FROM scheduled_emails
       WHERE status = 'pending' AND send_at <= datetime('now')
       ORDER BY send_at ASC
       LIMIT ?`,
      [cap],
      (err, rows) => {
        if (err) reject(err);
        else resolve(rows || []);
      }
    );
  });
};

module.exports.markScheduledEmailSent = ({ id, sent_at = null }) => {
  return new Promise((resolve, reject) => {
    const numericId = Number(id);
    if (!Number.isInteger(numericId) || numericId <= 0) {
      reject(new Error('Invalid scheduled email id'));
      return;
    }
    const sentAt = toSqliteDateTime(sent_at || new Date());
    db.run(
      `UPDATE scheduled_emails
       SET status = 'sent', sent_at = ?, last_error = NULL, updated_at = datetime('now')
       WHERE id = ?`,
      [sentAt, numericId],
      function (err) {
        if (err) reject(err);
        else resolve({ updated: this.changes || 0 });
      }
    );
  });
};

module.exports.markScheduledEmailFailed = ({ id, error, retryInMinutes = 5 }) => {
  return new Promise((resolve, reject) => {
    const numericId = Number(id);
    if (!Number.isInteger(numericId) || numericId <= 0) {
      reject(new Error('Invalid scheduled email id'));
      return;
    }
    const delay = Math.max(1, Number(retryInMinutes) || 5);
    const errorText = (error || '').toString().slice(0, 500);
    db.run(
      `UPDATE scheduled_emails
       SET attempt_count = attempt_count + 1,
           last_error = ?,
           send_at = datetime('now', ?),
           updated_at = datetime('now')
       WHERE id = ?`,
      [errorText, `+${delay} minutes`, numericId],
      function (err) {
        if (err) reject(err);
        else resolve({ updated: this.changes || 0 });
      }
    );
  });
};

module.exports.listScheduledEmails = ({ status = null, limit = 100, jobsheet_id = null, business_id = null, to = null, subject = null } = {}) => {
  return new Promise((resolve, reject) => {
    const where = [];
    const params = [];
    if (status) { where.push('status = ?'); params.push(status); }
    if (Number.isInteger(jobsheet_id)) { where.push('jobsheet_id = ?'); params.push(Number(jobsheet_id)); }
    if (Number.isInteger(business_id)) { where.push('business_id = ?'); params.push(Number(business_id)); }
    if (to) { where.push('to_address = ?'); params.push(String(to)); }
    if (subject) { where.push('subject = ?'); params.push(String(subject)); }
    const limitVal = Number(limit) || 100;
    params.push(limitVal);
    db.all(
      `SELECT * FROM scheduled_emails
       ${where.length ? 'WHERE ' + where.join(' AND ') : ''}
       ORDER BY send_at ASC
       LIMIT ?`,
      params,
      (err, rows) => {
        if (err) reject(err);
        else resolve(rows || []);
      }
    );
  });
};

module.exports.listPlannerActions = ({ business_id = null, jobsheet_id = null, status = null, action_key = null, after = null, before = null, limit = 500 } = {}) => {
  return new Promise((resolve, reject) => {
    const where = [];
    const params = [];
    if (Number.isInteger(business_id)) { where.push('business_id = ?'); params.push(Number(business_id)); }
    if (Number.isInteger(jobsheet_id)) { where.push('jobsheet_id = ?'); params.push(Number(jobsheet_id)); }
    if (status) { where.push('status = ?'); params.push(String(status)); }
    if (action_key) { where.push('action_key = ?'); params.push(String(action_key)); }
    const afterVal = toSqliteDateTime(after);
    if (afterVal) { where.push('scheduled_for >= ?'); params.push(afterVal); }
    const beforeVal = toSqliteDateTime(before);
    if (beforeVal) { where.push('scheduled_for <= ?'); params.push(beforeVal); }
    const limitVal = Number(limit) || 500;
    params.push(limitVal);
    db.all(
      `SELECT * FROM planner_actions
       ${where.length ? 'WHERE ' + where.join(' AND ') : ''}
       ORDER BY scheduled_for ASC
       LIMIT ?`,
      params,
      (err, rows) => {
        if (err) reject(err);
        else resolve(rows || []);
      }
    );
  });
};

module.exports.upsertPlannerAction = ({
  business_id,
  jobsheet_id,
  action_key,
  scheduled_for,
  status = null,
  completed_at = null,
  last_notified_at = null,
  last_email_at = null,
  last_error = null
}) => {
  return new Promise((resolve, reject) => {
    const bizId = Number(business_id);
    const jobId = Number(jobsheet_id);
    if (!Number.isInteger(bizId) || !Number.isInteger(jobId)) {
      reject(new Error('business_id and jobsheet_id are required'));
      return;
    }
    const key = String(action_key || '').trim();
    if (!key) { reject(new Error('action_key is required')); return; }
    const scheduledFor = toSqliteDateTime(scheduled_for);
    if (!scheduledFor) { reject(new Error('scheduled_for is invalid')); return; }
    const completedAt = toSqliteDateTime(completed_at);
    const notifiedAt = toSqliteDateTime(last_notified_at);
    const emailedAt = toSqliteDateTime(last_email_at);
    const statusVal = status ? String(status) : null;
    const errorText = last_error != null ? String(last_error).slice(0, 500) : null;
    db.run(
      `INSERT INTO planner_actions (business_id, jobsheet_id, action_key, scheduled_for, status, completed_at, last_notified_at, last_email_at, last_error, created_at, updated_at)
       VALUES (?, ?, ?, ?, COALESCE(?, 'pending'), ?, ?, ?, ?, datetime('now'), datetime('now'))
       ON CONFLICT(business_id, jobsheet_id, action_key, scheduled_for)
       DO UPDATE SET
         status = COALESCE(excluded.status, planner_actions.status),
         completed_at = COALESCE(excluded.completed_at, planner_actions.completed_at),
         last_notified_at = COALESCE(excluded.last_notified_at, planner_actions.last_notified_at),
         last_email_at = COALESCE(excluded.last_email_at, planner_actions.last_email_at),
         last_error = COALESCE(excluded.last_error, planner_actions.last_error),
         updated_at = datetime('now')`,
      [bizId, jobId, key, scheduledFor, statusVal, completedAt, notifiedAt, emailedAt, errorText],
      function (err) {
        if (err) reject(err);
        else resolve({ ok: true, id: this.lastID });
      }
    );
  });
};

module.exports.updatePlannerActionById = ({ action_id, status = null, completed_at = null, last_notified_at = null, last_email_at = null, last_error = null } = {}) => {
  return new Promise((resolve, reject) => {
    const id = Number(action_id);
    if (!Number.isInteger(id) || id <= 0) {
      reject(new Error('Invalid action id'));
      return;
    }
    const statusVal = status ? String(status) : null;
    const completedAt = toSqliteDateTime(completed_at);
    const notifiedAt = toSqliteDateTime(last_notified_at);
    const emailedAt = toSqliteDateTime(last_email_at);
    const errorText = last_error != null ? String(last_error).slice(0, 500) : null;
    db.run(
      `UPDATE planner_actions
       SET status = COALESCE(?, status),
           completed_at = COALESCE(?, completed_at),
           last_notified_at = COALESCE(?, last_notified_at),
           last_email_at = COALESCE(?, last_email_at),
           last_error = COALESCE(?, last_error),
           updated_at = datetime('now')
       WHERE action_id = ?`,
      [statusVal, completedAt, notifiedAt, emailedAt, errorText, id],
      function (err) {
        if (err) reject(err);
        else resolve({ updated: this.changes || 0 });
      }
    );
  });
};

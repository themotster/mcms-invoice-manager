const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const path = require('path');

const settings = JSON.parse(fs.readFileSync(path.join(__dirname, 'settings.json')));
const db = new sqlite3.Database(settings.db_path);
const mergeFieldDefaults = require('./config/mergeFields.json');

const DEFAULT_BUSINESS_ID = 1;
const DEFAULT_BUSINESSES = [
  {
    id: 1,
    business_name: 'Motti Cohen Music Services',
    last_invoice_number: 704,
    last_quote_number: 0,
    save_path: '/Users/Shared/Invoices/MCMS',
    invoice_template_path: null,
    quote_template_path: null,
    contract_template_path: null,
    gig_sheet_template_path: null
  },
  {
    id: 2,
    business_name: 'AhMen A Cappella Ltd',
    last_invoice_number: 882,
    last_quote_number: 0,
    save_path: '/Users/Shared/Invoices/AhMen',
    invoice_template_path: null,
    quote_template_path: null,
    contract_template_path: null,
    gig_sheet_template_path: null
  }
];

const DEFAULT_DOCUMENT_DEFINITIONS = [
  {
    key: 'workbook',
    doc_type: 'workbook',
    label: 'Excel Workbook',
    description: 'Excel workbook populated with all jobsheet details.',
    file_suffix: '',
    invoice_variant: null,
    requires_total: 1,
    is_primary: 1,
    sort_order: 0,
    locked: 1
  },
  {
    key: 'quote',
    doc_type: 'quote',
    label: 'Quote',
    description: 'Quote document with pricing totals.',
    file_suffix: ' - Quote',
    invoice_variant: null,
    requires_total: 1,
    is_primary: 0,
    sort_order: 1,
    locked: 1
  },
  {
    key: 'contract',
    doc_type: 'contract',
    label: 'Contract',
    description: 'Contract ready for signatures.',
    file_suffix: ' - Contract',
    invoice_variant: null,
    requires_total: 0,
    is_primary: 0,
    sort_order: 2,
    locked: 1
  },
  {
    key: 'invoice_deposit',
    doc_type: 'invoice',
    label: 'Invoice – Deposit',
    description: 'Deposit invoice for the booking.',
    file_suffix: ' - Deposit',
    invoice_variant: 'deposit',
    requires_total: 1,
    is_primary: 0,
    sort_order: 3,
    locked: 1
  },
  {
    key: 'invoice_balance',
    doc_type: 'invoice',
    label: 'Invoice – Balance',
    description: 'Balance invoice for the booking.',
    file_suffix: ' - Balance',
    invoice_variant: 'balance',
    requires_total: 1,
    is_primary: 0,
    sort_order: 4,
    locked: 1
  }
];

const BUSINESS_SETTINGS_MUTABLE_FIELDS = new Set([
  'save_path',
  'invoice_template_path',
  'quote_template_path',
  'contract_template_path',
  'gig_sheet_template_path',
  'last_invoice_number',
  'last_quote_number'
]);

const AHMEN_JOBSHEET_FIELDS = [
  'business_id',
  'status',
  'client_name',
  'client_email',
  'client_phone',
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
  'venue_address1',
  'venue_address2',
  'venue_address3',
  'venue_town',
  'venue_postcode',
  'venue_same_as_client',
  'ahmen_fee',
  'production_fees',
  'deposit_amount',
  'balance_amount',
  'balance_due_date',
  'balance_reminder_date',
  'service_types',
  'specialist_singers',
  'notes',
  'pricing_service_id',
  'pricing_selected_singers',
  'pricing_custom_fees',
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
  'venue_same_as_client'
]);

const AHMEN_JOBSHEET_INTEGER_FIELDS = new Set([
  'venue_id'
]);

const AHMEN_JOBSHEET_STATUS_VALUES = new Set(['enquiry', 'quoted', 'confirmed', 'completed']);

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
      save_path TEXT NOT NULL,
      invoice_template_path TEXT,
      quote_template_path TEXT,
      contract_template_path TEXT,
      gig_sheet_template_path TEXT
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

    db.run(`CREATE TABLE IF NOT EXISTS document_definitions (
      definition_id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_id INTEGER NOT NULL,
      key TEXT NOT NULL,
      doc_type TEXT NOT NULL,
      label TEXT NOT NULL,
      description TEXT,
      file_suffix TEXT,
      invoice_variant TEXT,
      template_path TEXT,
      requires_total INTEGER DEFAULT 0,
      is_primary INTEGER DEFAULT 0,
      is_active INTEGER DEFAULT 1,
      is_locked INTEGER DEFAULT 0,
      sort_order INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (business_id) REFERENCES business_settings(id),
      UNIQUE (business_id, key)
    )`);

    db.run(`CREATE INDEX IF NOT EXISTS idx_document_definitions_business_sort
      ON document_definitions (business_id, sort_order)`);

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
      deposit_amount REAL,
      balance_amount REAL,
      balance_due_date TEXT,
      balance_reminder_date TEXT,
      service_types TEXT,
      specialist_singers TEXT,
      notes TEXT,
      pricing_service_id TEXT,
      pricing_selected_singers TEXT,
      pricing_custom_fees TEXT,
      pricing_discount REAL,
      pricing_discount_type TEXT,
      pricing_discount_value REAL,
      pricing_production_items TEXT,
      pricing_production_subtotal REAL,
      pricing_production_discount TEXT,
      pricing_production_discount_type TEXT,
      pricing_production_discount_value REAL,
      pricing_production_total REAL,
      pricing_total REAL,
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
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN caterer_name TEXT', logDuplicateColumn);

    db.run('ALTER TABLE business_settings ADD COLUMN invoice_template_path TEXT', logDuplicateColumn);
    db.run('ALTER TABLE business_settings ADD COLUMN quote_template_path TEXT', logDuplicateColumn);
    db.run('ALTER TABLE business_settings ADD COLUMN contract_template_path TEXT', logDuplicateColumn);
    db.run('ALTER TABLE business_settings ADD COLUMN gig_sheet_template_path TEXT', logDuplicateColumn);
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
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_discount REAL', logDuplicateColumn);
    db.run('ALTER TABLE ahmen_jobsheets ADD COLUMN pricing_total REAL', logDuplicateColumn);
    db.run(`UPDATE ahmen_jobsheets SET status='enquiry' WHERE status IS NULL OR status='' OR status='draft'`, err => {
      if (err) console.error('Failed to normalize jobsheet status:', err);
    });

    seedBusinesses();
    syncLegacyBusinessesTable();
    seedMergeFieldDefaults();
    seedDocumentDefinitions();
  });
}

function seedBusinesses() {
  const allowedNames = DEFAULT_BUSINESSES.map(b => b.business_name);

  DEFAULT_BUSINESSES.forEach(business => {
    db.run(
      `INSERT OR IGNORE INTO business_settings (
        id, business_name, last_invoice_number, last_quote_number, save_path,
        invoice_template_path, quote_template_path, contract_template_path, gig_sheet_template_path
      )
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        business.id,
        business.business_name,
        business.last_invoice_number,
        business.last_quote_number,
        business.save_path,
        business.invoice_template_path,
        business.quote_template_path,
        business.contract_template_path,
        business.gig_sheet_template_path
      ]
    );

    db.run(
      `UPDATE business_settings
       SET business_name = ?, last_invoice_number = ?, last_quote_number = ?,
           save_path = CASE WHEN save_path IS NULL OR save_path = '' THEN ? ELSE save_path END,
           invoice_template_path = COALESCE(?, invoice_template_path),
           quote_template_path = COALESCE(?, quote_template_path),
           contract_template_path = COALESCE(?, contract_template_path),
           gig_sheet_template_path = COALESCE(?, gig_sheet_template_path)
       WHERE id = ?`,
      [
        business.business_name,
        business.last_invoice_number,
        business.last_quote_number,
        business.save_path,
        business.invoice_template_path,
        business.quote_template_path,
        business.contract_template_path,
        business.gig_sheet_template_path,
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

function resolveDefinitionTemplatePath(business, definitionKey) {
  if (!business) return null;
  switch (definitionKey) {
    case 'workbook':
      return business.gig_sheet_template_path || business.invoice_template_path || null;
    case 'quote':
      return business.quote_template_path || null;
    case 'contract':
      return business.contract_template_path || null;
    case 'invoice_deposit':
    case 'invoice_balance':
      return business.invoice_template_path || null;
    default:
      return null;
  }
}

function seedDocumentDefinitions() {
  db.all(`SELECT id, invoice_template_path, quote_template_path, contract_template_path, gig_sheet_template_path FROM business_settings`, (err, businesses) => {
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

          DEFAULT_DOCUMENT_DEFINITIONS.forEach(definition => {
            const templatePath = resolveDefinitionTemplatePath(business, definition.key);

            db.run(
              `INSERT OR IGNORE INTO document_definitions (
                business_id,
                key,
                doc_type,
                label,
                description,
                file_suffix,
                invoice_variant,
                template_path,
                requires_total,
                is_primary,
                is_active,
                is_locked,
                sort_order,
                created_at,
                updated_at
              ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, datetime('now'), datetime('now'))`,
              [
                businessId,
                definition.key,
                definition.doc_type,
                definition.label,
                definition.description || null,
                definition.file_suffix || null,
                definition.invoice_variant || null,
                templatePath || null,
                definition.requires_total ? 1 : 0,
                definition.is_primary ? 1 : 0,
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
                  file_suffix = ?,
                  invoice_variant = ?,
                  requires_total = ?,
                  is_primary = ?,
                  is_locked = CASE WHEN is_locked = 1 THEN 1 ELSE ? END,
                  sort_order = CASE WHEN is_locked = 1 THEN ? ELSE sort_order END,
                  template_path = CASE WHEN ? IS NOT NULL AND (template_path IS NULL OR template_path = '') THEN ? ELSE template_path END
              WHERE business_id = ? AND key = ?`;

            db.run(
              updateSql,
              [
                definition.doc_type,
                definition.label,
                definition.description || null,
                definition.file_suffix || null,
                definition.invoice_variant || null,
                definition.requires_total ? 1 : 0,
                definition.is_primary ? 1 : 0,
                definition.locked ? 1 : 0,
                definition.sort_order,
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
    });
  });
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
  return {
    ...row,
    requires_total: row.requires_total ? 1 : 0,
    is_primary: row.is_primary ? 1 : 0,
    is_active: row.is_active ? 1 : 0,
    is_locked: row.is_locked ? 1 : 0,
    sort_order: Number.isFinite(row.sort_order) ? Number(row.sort_order) : 0
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
  return {
    key: (definition.key || '').trim(),
    doc_type: (definition.doc_type || '').trim(),
    label: (definition.label || '').trim(),
    description: definition.description != null && definition.description !== '' ? String(definition.description) : null,
    file_suffix: definition.file_suffix != null && definition.file_suffix !== '' ? String(definition.file_suffix) : null,
    invoice_variant: definition.invoice_variant != null && definition.invoice_variant !== '' ? String(definition.invoice_variant) : null,
    template_path: definition.template_path != null && definition.template_path !== '' ? String(definition.template_path) : null,
    requires_total: definition.requires_total ? 1 : 0,
    is_primary: definition.is_primary ? 1 : 0,
    is_active: definition.is_active === 0 ? 0 : 1,
    is_locked: definition.is_locked ? 1 : 0,
    sort_order: Number.isFinite(definition.sort_order) ? Number(definition.sort_order) : null
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
           file_suffix,
           invoice_variant,
           template_path,
           requires_total,
           is_primary,
           is_active,
           is_locked,
           sort_order,
           created_at,
           updated_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(business_id, key) DO UPDATE SET
           doc_type = excluded.doc_type,
           label = excluded.label,
           description = excluded.description,
           file_suffix = excluded.file_suffix,
           invoice_variant = excluded.invoice_variant,
           template_path = excluded.template_path,
           requires_total = excluded.requires_total,
           is_primary = excluded.is_primary,
           is_active = excluded.is_active,
           sort_order = excluded.sort_order,
           updated_at = excluded.updated_at,
           is_locked = CASE WHEN document_definitions.is_locked = 1 THEN 1 ELSE excluded.is_locked END`,
        [
          id,
          payload.key,
          payload.doc_type,
          payload.label,
          payload.description,
          payload.file_suffix,
          payload.invoice_variant,
          payload.template_path,
          payload.requires_total,
          payload.is_primary,
          payload.is_active,
          payload.is_locked,
          orderValue,
          now,
          now
        ],
        function (err) {
          if (err) {
            reject(err);
          } else {
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
      ? 'SELECT definition_id, is_locked FROM document_definitions WHERE business_id = ? AND definition_id = ?'
      : 'SELECT definition_id, is_locked FROM document_definitions WHERE business_id = ? AND key = ?';

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

      db.run(
        "UPDATE document_definitions SET is_active = 0, updated_at = datetime('now') WHERE definition_id = ?",
        [row.definition_id],
        function (updateErr) {
          if (updateErr) reject(updateErr);
          else resolve({ removed: this.changes });
        }
      );
    });
  });
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

function applyTemplatePathUpdates(businessId, updates = {}) {
  const id = Number(businessId);
  if (!Number.isInteger(id)) return;

  const templateMap = {
    invoice_template_path: ['invoice_deposit', 'invoice_balance'],
    quote_template_path: ['quote'],
    contract_template_path: ['contract'],
    gig_sheet_template_path: ['workbook']
  };

  Object.entries(templateMap).forEach(([field, keys]) => {
    if (!(field in updates)) return;
    const value = updates[field] == null || updates[field] === '' ? null : updates[field];
    keys.forEach(key => {
      db.run(
        "UPDATE document_definitions SET template_path = ?, updated_at = datetime('now') WHERE business_id = ? AND key = ?",
        [value, id, key],
        err => {
          if (err) console.error('Failed to sync template path for definition', key, err);
        }
      );
    });
  });
}

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
              applyTemplatePathUpdates(id, updates);
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

function deleteAhmenJobsheet(jobsheetId) {
  return new Promise((resolve, reject) => {
    const id = Number(jobsheetId);
    if (!Number.isInteger(id)) {
      reject(new Error('Invalid jobsheet id'));
      return;
    }

    db.run(
      'DELETE FROM ahmen_jobsheets WHERE jobsheet_id = ?',
      [id],
      function (err) {
        if (err) reject(err);
        else resolve(this.changes);
      }
    );
  });
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
           def.file_suffix AS definition_file_suffix,
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
      if (data.file_path !== undefined) {
        updates.push('file_path = ?');
        params.push(data.file_path);
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

  getBusinessById: (businessId) => {
    return new Promise((resolve, reject) => {
      db.get(
        `SELECT * FROM business_settings WHERE id = ?`,
        [businessId],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });
  },

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
  deleteAhmenJobsheet,
  getAhmenVenues,
  saveAhmenVenue,
  deleteAhmenVenue,
  getMergeFields,
  getMergeFieldBindingsByTemplate,
  saveMergeField,
  deleteMergeField
};

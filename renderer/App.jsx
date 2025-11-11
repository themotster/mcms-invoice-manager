import React, { Fragment, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import TemplatesManager from './components/TemplatesManager';
import WysiwygEditor from './components/WysiwygEditor.jsx';
import { ExcelTemplateEditor } from './components/InvoiceCanvasEditor.jsx';
import ToastOverlay from './components/ToastOverlay';
import MailComposer from './components/MailComposer';
import { normalizeVenues, buildVenueDraft } from './helpers/venues';
import {
  normalizeProductionItems,
  calculateProductionItemTotal,
  calculateProductionTotal,
  calculateDiscountValue
} from './helpers/pricing';

const AHMEN_NUMERIC_FIELDS = new Set([
  'ahmen_fee',
  'production_fees',
  'deposit_amount',
  'balance_amount',
  'pricing_discount',
  'pricing_total',
  'pricing_production_subtotal',
  'pricing_production_total',
  'pricing_discount_value',
  'pricing_production_discount_value'
]);

const AHMEN_BOOLEAN_FIELDS = new Set(['venue_same_as_client']);

const STATUS_OPTIONS = [
  { value: 'enquiry', label: 'Enquiry' },
  { value: 'quoted', label: 'Quoted' },
  { value: 'contracting', label: 'Contracting' },
  { value: 'confirmed', label: 'Confirmed' },
  { value: 'completed', label: 'Completed' }
];

const DOCUMENT_TYPE_LABELS = {
  invoice: 'Invoice',
  quote: 'Quote',
  contract: 'Contract',
  workbook: 'Excel Workbook',
  pdf_export: 'PDF Export'
};

const DOC_TYPE_META = {
  invoice: {
    label: DOCUMENT_TYPE_LABELS.invoice,
    filters: [{ name: 'Excel workbooks', extensions: ['xlsx'] }],
    supportsNormalize: true
  },
  quote: {
    label: DOCUMENT_TYPE_LABELS.quote,
    filters: [{ name: 'Excel workbooks', extensions: ['xlsx'] }],
    supportsNormalize: true
  },
  contract: {
    label: DOCUMENT_TYPE_LABELS.contract,
    filters: [{ name: 'Word documents', extensions: ['docx'] }],
    supportsNormalize: false
  },
  workbook: {
    label: DOCUMENT_TYPE_LABELS.workbook,
    filters: [{ name: 'Excel workbooks', extensions: ['xlsx'] }],
    supportsNormalize: true
  }
};

const DOCUMENT_TYPE_OPTIONS = Object.entries(DOC_TYPE_META).map(([value, meta]) => ({
  value,
  label: meta.label
}));

const BOOKING_PACK_DEFINITION_KEYS = new Set(['booking_schedule', 't_cs', 'invoice_deposit']);

const DOCUMENT_GROUP_OPTIONS = [
  { value: 'none', label: 'All Documents' },
  { value: 'doc_type', label: 'Document Type' },
  { value: 'client', label: 'Client' },
  { value: 'event_date', label: 'Event Date' }
];

const DOCUMENT_CARD_TONES = {
  workbook: {
    outerBorder: 'border-teal-200',
    outerBg: 'rgba(209,250,229,0.85)',
    innerBorder: 'border-teal-200'
  },
  quote: {
    outerBorder: 'border-sky-200',
    outerBg: 'rgba(224,242,254,0.85)',
    innerBorder: 'border-sky-200'
  },
  contract: {
    outerBorder: 'border-violet-200',
    outerBg: 'rgba(237,233,254,0.85)',
    innerBorder: 'border-violet-200'
  },
  invoice: {
    outerBorder: 'border-amber-200',
    outerBg: 'rgba(254,243,199,0.85)',
    innerBorder: 'border-amber-200'
  },
  client_data: {
    outerBorder: 'border-lime-200',
    outerBg: 'rgba(236,252,203,0.85)',
    innerBorder: 'border-lime-200'
  },
  default: {
    outerBorder: 'border-slate-200',
    outerBg: 'rgba(248,250,252,0.9)',
    innerBorder: 'border-slate-200'
  }
};

const DOCUMENT_COLUMNS = [
  { key: 'document', label: 'Document', align: 'left', always: true },
  { key: 'client', label: 'Client / Event', align: 'left' },
  { key: 'event_date', label: 'Event Date', align: 'left' },
  { key: 'created', label: 'Created', align: 'left' },
  { key: 'amount', label: 'Amount', align: 'right' },
  { key: 'actions', label: 'Actions', align: 'right', always: true }
];

const DOCUMENT_FEATURES_ENABLED = true;
const DOCUMENT_GENERATION_ENABLED = true;
const HARD_LOCKED_DEFINITION_KEYS = new Set(['workbook']);

function getDocumentIcon(docType) {
  switch ((docType || '').toLowerCase()) {
    case 'invoice':
      return '🧾';
    case 'quote':
      return '💼';
    case 'contract':
      return '🖋️';
    case 'workbook':
      return '📊';
    case 'pdf_export':
      return '🖨️';
    default:
      return '📄';
  }
}

const WORKSPACE_ICON_MAP = {
  jobsheets: '🗂️',
  documents: '🗃️',
  invoices: '🧾',
  templates: '📁',
  settings: '⚙️'
};

const WORKSPACE_SECTIONS = [
  { key: 'jobsheets', label: 'Jobsheets', description: 'Bookings and statuses', icon: WORKSPACE_ICON_MAP.jobsheets },
  { key: 'documents', label: 'Documents', description: 'Browse and manage files', icon: WORKSPACE_ICON_MAP.documents },
  { key: 'invoices', label: 'Invoice Log', description: 'Issued invoices and status', icon: WORKSPACE_ICON_MAP.invoices },
  { key: 'templates', label: 'Templates', description: 'Manage document templates', icon: WORKSPACE_ICON_MAP.templates },
  { key: 'settings', label: 'Settings', description: 'Business preferences', icon: WORKSPACE_ICON_MAP.settings }
];

const WORKSPACE_SECTION_STORAGE_KEY = 'invoiceMaster:workspaceSection';
const DOCUMENT_COLUMNS_STORAGE_KEY = 'invoiceMaster:documentsColumns';
const DOCUMENT_TREE_COLLAPSE_KEY = 'invoiceMaster:documentTreeCollapsed';
const DEFAULT_DOCUMENT_COLUMNS_STATE = DOCUMENT_COLUMNS.reduce((acc, column) => {
  if (!column.always) {
    acc[column.key] = true;
  }
  return acc;
}, {});

function slugifyDefinitionKey(value) {
  return (value || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 60);
}

function createDefinitionDraft(overrides = {}) {
  return {
    key: '',
    label: '',
    doc_type: 'invoice',
    description: '',
    invoice_variant: '',
    template_path: '',
    is_primary: 0,
    is_active: 1,
    is_locked: 0,
    sort_order: null,
    ...overrides
  };
}

const STATUS_STYLES = {
  enquiry: 'bg-yellow-100 text-yellow-800 border border-yellow-200',
  quoted: 'bg-blue-100 text-blue-800 border border-blue-200',
  contracting: 'bg-indigo-100 text-indigo-800 border border-indigo-200',
  confirmed: 'bg-green-100 text-green-800 border border-green-200',
  completed: 'bg-gray-200 text-gray-700 border border-gray-300'
};

const STATUS_ROW_CLASSES = {
  enquiry: 'bg-yellow-100',
  quoted: 'bg-blue-100',
  contracting: 'bg-indigo-100',
  confirmed: 'bg-green-100',
  completed: 'bg-gray-200'
};

const ACTIVE_STATUS_ROW_CLASSES = {
  enquiry: 'bg-yellow-400',
  quoted: 'bg-blue-400',
  contracting: 'bg-indigo-400',
  confirmed: 'bg-green-400',
  completed: 'bg-gray-500'
};

const STATUS_ORDER = STATUS_OPTIONS.reduce((acc, option, index) => {
  acc[option.value] = index;
  return acc;
}, {});

const STATUS_DOT_CLASSES = {
  enquiry: 'bg-yellow-400',
  quoted: 'bg-blue-400',
  contracting: 'bg-indigo-400',
  confirmed: 'bg-green-500',
  completed: 'bg-slate-400'
};

const LAST_BUSINESS_STORAGE_KEY = 'invoiceMaster:lastBusinessId';

function readLastBusinessId() {
  try {
    return window.localStorage.getItem(LAST_BUSINESS_STORAGE_KEY);
  } catch (err) {
    console.warn('Unable to read last business id', err);
    return null;
  }
}

function storeLastBusinessId(id) {
  try {
    if (id) {
      window.localStorage.setItem(LAST_BUSINESS_STORAGE_KEY, String(id));
    } else {
      window.localStorage.removeItem(LAST_BUSINESS_STORAGE_KEY);
    }
  } catch (err) {
    console.warn('Unable to persist last business id', err);
  }
}

function normalizeStatus(value) {
  if (!value) return '';
  if (typeof value === 'string') return value.toLowerCase();
  return String(value).toLowerCase();
}

const JOBSHEET_COLUMNS = [
  { key: 'client_name', label: 'Client', sortable: true, align: 'left' },
  { key: 'event_type', label: 'Event', sortable: true, align: 'left' },
  { key: 'event_date', label: 'Event Date', sortable: true, align: 'left' },
  { key: 'venue_name', label: 'Venue', sortable: true, align: 'left' },
  { key: 'status', label: 'Status', sortable: true, align: 'center' },
  { key: 'ahmen_fee', label: 'Fee', sortable: true, align: 'right' },
  { key: 'actions', label: '', sortable: false, align: 'right' }
];

const JOBSHEET_GRID_TEMPLATE = 'minmax(0,1.4fr) minmax(0,1fr) minmax(0,1.15fr) minmax(0,1.4fr) minmax(0,0.9fr) minmax(0,1fr) auto';

function getAlignmentClasses(alignment) {
  if (alignment === 'right') return 'justify-end text-right';
  if (alignment === 'center') return 'justify-center text-center';
  return 'justify-start text-left';
}

const DEFAULT_JOBSHEET = (businessId) => ({
  business_id: businessId,
  status: 'enquiry',
  client_name: '',
  client_email: '',
  client_phone: '',
  client_address1: '',
  client_address2: '',
  client_address3: '',
  client_town: '',
  client_postcode: '',
  event_type: '',
  event_date: '',
  event_start: '',
  event_end: '',
  venue_id: null,
  venue_same_as_client: false,
  venue_name: '',
  venue_address1: '',
  venue_address2: '',
  venue_address3: '',
  venue_town: '',
  venue_postcode: '',
  ahmen_fee: '',
  production_fees: '',
  deposit_amount: '',
  balance_amount: '',
  balance_due_date: '',
  balance_reminder_date: '',
  service_types: '',
  notes: '',
  pricing_service_id: '',
  pricing_selected_singers: [],
  pricing_custom_fees: '',
  pricing_discount: '',
  pricing_discount_type: 'amount',
  pricing_discount_value: '',
  pricing_production_items: [],
  pricing_production_subtotal: '',
  pricing_production_discount: '',
  pricing_production_discount_type: 'amount',
  pricing_production_discount_value: '',
  pricing_production_total: '',
  pricing_total: ''
});

const CATEGORY_TO_GROUP_KEY = {
  client: 'client',
  event: 'event',
  venue: 'venue',
  financial: 'billing',
  services: 'services'
};

const FIELD_META = {
  status: {
    label: 'Status',
    component: 'statusSelect',
    options: STATUS_OPTIONS,
    always: true
  },
  saved_venue: {
    label: 'Saved Venue',
    component: 'savedVenueSelector',
    always: true
  },
  venue_same_as_client: {
    label: 'Use client address (private residence)',
    type: 'checkbox',
    hint: 'Copies the client address and does not save the venue to the shared directory.',
    always: true
  },
  pricing_panel: {
    component: 'pricingPanel',
    always: true
  },
  production_panel: {
    component: 'productionPanel',
    always: true
  },
  documents_panel: {
    component: 'documentsPanel',
    always: true
  },
  gig_info_panel: {
    component: 'gigInfoPanel',
    always: true
  },
  notes: {
    label: 'Internal Notes',
    type: 'textarea',
    rows: 3,
    always: true,
    fallback: true
  },
  client_name: { fallback: true },
  client_email: { type: 'email', fallback: true },
  client_phone: { type: 'tel', fallback: true },
  client_address1: { fallback: true },
  client_address2: { fallback: true },
  client_address3: { fallback: true },
  client_town: { fallback: true },
  client_postcode: { fallback: true },
  event_type: { fallback: true },
  event_date: { type: 'date', fallback: true },
  event_start: { type: 'time', fallback: true },
  event_end: { type: 'time', fallback: true },
  venue_name: { fallback: true },
  venue_address1: { fallback: true },
  venue_address2: { fallback: true },
  venue_address3: { fallback: true },
  venue_town: { fallback: true },
  venue_postcode: { fallback: true },
  caterer_name: {},
  service_types: { type: 'textarea', rows: 2, fallback: true },
  specialist_singers: { type: 'textarea', rows: 2 },
  ahmen_fee: {
    label: 'AhMen Fee (£)',
    type: 'number',
    step: '0.01',
    readOnly: true,
    hint: 'Singer fees after discount.',
    always: true,
    fallback: true
  },
  production_fees: {
    label: 'Sound / AV / Production (£)',
    type: 'number',
    step: '0.01',
    readOnly: true,
    always: true,
    fallback: true
  },
  extra_fees: {
    label: 'Extra Fees',
    type: 'number',
    step: '0.01'
  },
  total_amount: {
    label: 'Total Amount',
    type: 'number',
    step: '0.01',
    readOnly: true
  },
  deposit_amount: {
    label: 'Deposit (£)',
    type: 'number',
    step: '0.01',
    readOnly: true,
    hint: 'Automatically 30% of AhMen fee.',
    always: true,
    fallback: true
  },
  balance_amount: {
    label: 'Balance (£)',
    type: 'number',
    step: '0.01',
    readOnly: true,
    hint: 'Remaining balance after deposit (70%).',
    always: true,
    fallback: true
  },
  balance_due_date: {
    label: 'Balance Due Date',
    type: 'date',
    readOnly: true,
    hint: 'Automatically 10 days before the event.',
    always: true,
    fallback: true
  },
  balance_reminder_date: {
    label: 'Balance Reminder Date',
    type: 'date',
    readOnly: true,
    hint: 'Automatically 20 days before the event.',
    always: true,
    fallback: true
  }
};

const GROUP_CONFIG = {
  client: {
    title: 'Client Details',
    description: 'Captured during the initial enquiry.',
    category: 'client',
    order: [
      'client_name',
      'client_email',
      'client_phone',
      'client_address1',
      'client_address2',
      'client_address3',
      'client_town',
      'client_postcode'
    ],
    prepend: ['status'],
    defaultOpen: true
  },
  event: {
    title: 'Event Details',
    description: 'What, when, and how the event will run.',
    category: 'event',
    order: ['event_type', 'event_date', 'event_start', 'event_end']
  },
  venue: {
    title: 'Venue Details',
    description: 'Where your team will be performing and saved venue options.',
    category: 'venue',
    order: [
      'venue_name',
      'venue_address1',
      'venue_address2',
      'venue_address3',
      'venue_town',
      'venue_postcode',
      'caterer_name'
    ],
    prepend: ['saved_venue', 'venue_same_as_client']
  },
  pricing: {
    title: 'Pricing & Personnel',
    description: 'Select singers and configure fees for the booking.',
    staticOnly: true,
    fields: ['pricing_panel']
  },
  production: {
    title: 'Production & Services',
    description: 'Manage external suppliers, markup, and related discounts.',
    staticOnly: true,
    fields: ['production_panel']
  },
  billing: {
    title: 'Invoicing Details',
    description: 'Invoicing breakdown that feeds quotes and invoices.',
    category: 'financial',
    order: [
      'ahmen_fee',
      'production_fees',
      'extra_fees',
      'total_amount',
      'deposit_amount',
      'balance_amount',
      'balance_due_date',
      'balance_reminder_date'
    ]
  },
  services: {
    title: 'Services & Notes',
    description: 'Additional requirements and context for the booking.',
    category: 'services',
    order: ['service_types', 'specialist_singers'],
    append: ['notes']
  },
  documents: {
    title: 'Documents',
    description: 'Generate Excel outputs and manage PDFs.',
    staticOnly: true,
    fields: ['documents_panel']
  },
  gig_info: {
    title: 'Gig Info',
    description: 'Fill in and generate a singer-facing gig info PDF.',
    staticOnly: true,
    fields: ['gig_info_panel']
  }
};

const GROUP_ICON_MAP = {
  client: '👤',
  event: '🎉',
  venue: '📍',
  pricing: '🎶',
  production: '🎛️',
  billing: '💷',
  services: '📝',
  documents: '🗂️'
};

// Preserve inline mail composer state across transient unmounts (e.g. template refresh)
const COMPOSER_STORAGE_PREFIX = 'invoiceMaster:composerState:';

const loadComposerState = (key) => {
  if (!key || typeof window === 'undefined' || !window.sessionStorage) return null;
  try {
    const raw = window.sessionStorage.getItem(`${COMPOSER_STORAGE_PREFIX}${key}`);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (_) {
    return null;
  }
};

const persistComposerState = (key, value) => {
  if (!key || typeof window === 'undefined' || !window.sessionStorage) return;
  try {
    window.sessionStorage.setItem(`${COMPOSER_STORAGE_PREFIX}${key}`, JSON.stringify(value));
  } catch (_) {}
};

const clearComposerState = (key) => {
  if (!key || typeof window === 'undefined' || !window.sessionStorage) return;
  try {
    window.sessionStorage.removeItem(`${COMPOSER_STORAGE_PREFIX}${key}`);
  } catch (_) {}
};

const MAIL_TOKEN_REGEX = /{{\s*([a-zA-Z0-9_.-]+)(?:\|([^}]+))?\s*}}/g;

function buildMailTokenMap(snapshot = {}) {
  const js = snapshot || {};
  const fmtDate = (value) => {
    if (!value) return '';
    const str = String(value);
    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      try {
        const d = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
        return d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });
      } catch (_) {
        return str;
      }
    }
    return str;
  };

  const firstName = (() => {
    const raw = String(js.client_name || '').trim();
    if (!raw) return '';
    const parts = raw.split(/\s+/);
    return parts[0] || '';
  })();

  return {
    client_name: js.client_name || '',
    client_first_name: firstName,
    client_email: js.client_email || '',
    event_type: js.event_type || '',
    event_date: fmtDate(js.event_date || ''),
    balance_due_date: fmtDate(js.balance_due_date || ''),
    balance_reminder_date: fmtDate(js.balance_reminder_date || ''),
    today: fmtDate(new Date().toISOString().slice(0, 10))
  };
}

function renderMailTemplate(template, tokenMap = {}) {
  if (!template) return '';
  return String(template).replace(MAIL_TOKEN_REGEX, (_match, key, fallback) => {
    const normalizedKey = String(key || '').trim().toLowerCase();
    const value = tokenMap[normalizedKey];
    if (value != null && value !== '') return String(value);
    return fallback != null ? String(fallback) : '';
  });
}

async function resolveTemplateSubjectBody(api, businessId, jobsheetSnapshot, key) {
  try {
    const [templates, defaults] = await Promise.all([
      api?.getMailTemplates?.({ businessId }),
      api?.getDefaultMailTemplates?.({ businessId })
    ]);
    const def = (defaults && defaults[key]) || {};
    const custom = (templates && templates[key]) || {};
    const tpl = { ...def, ...custom };
    if (!tpl || (tpl.subject == null && tpl.body == null)) return { subject: '', body: '' };
    const tokenMap = buildMailTokenMap(jobsheetSnapshot || {});
    return {
      subject: renderMailTemplate(tpl.subject || '', tokenMap),
      body: renderMailTemplate(tpl.body || '', tokenMap)
    };
  } catch (_err) {
    return { subject: '', body: '' };
  }
}

function appendSignatureHtml(bodyHtml, signatureHtml) {
  const trimmedBody = (bodyHtml || '').trim();
  if (!signatureHtml) return trimmedBody;
  if (!trimmedBody) return signatureHtml;
  if (/(<br\s*\/?>|<\/p>)$/i.test(trimmedBody)) {
    return `${trimmedBody}${signatureHtml}`;
  }
  return `${trimmedBody}<br><br>${signatureHtml}`;
}

function startCaseKey(key) {
  if (!key) return '';
  return key
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\w\S*/g, word => word.charAt(0).toUpperCase() + word.slice(1));
}

function parseDateValue(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Number.isNaN(value.valueOf()) ? null : value;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const sqlDateTimePattern = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
    if (sqlDateTimePattern.test(trimmed)) {
      const parsed = new Date(`${trimmed.replace(' ', 'T')}Z`);
      return Number.isNaN(parsed.valueOf()) ? null : parsed;
    }
    const parsed = new Date(trimmed);
    return Number.isNaN(parsed.valueOf()) ? null : parsed;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

function formatCompactDate(value) {
  const date = parseDateValue(value);
  if (!date) return '—';
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: '2-digit'
  });
}

function IndeterminateCheckbox({ checked, indeterminate, className = '', ...props }) {
  const ref = useRef(null);
  useEffect(() => {
    if (ref.current) {
      ref.current.indeterminate = Boolean(indeterminate);
    }
  }, [indeterminate, checked]);
  return (
    <input
      type="checkbox"
      ref={ref}
      checked={checked}
      className={`h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 ${className}`}
      {...props}
    />
  );
}

function buildFieldConfig(fieldKey, registryField) {
  const meta = FIELD_META[fieldKey] || {};
  const label = registryField?.label || meta.label || startCaseKey(fieldKey);
  const hint = meta.hint !== undefined ? meta.hint : registryField?.description;
  const config = {
    name: fieldKey,
    label
  };

  if (meta.component) {
    config.component = meta.component;
    if (meta.options) config.options = meta.options;
  } else {
    config.type = meta.type || 'text';
  }

  if (meta.rows != null) config.rows = meta.rows;
  if (meta.readOnly != null) config.readOnly = meta.readOnly;
  if (meta.step != null) config.step = meta.step;
  if (hint) config.hint = hint;

  return config;
}

function buildJobsheetGroups(mergeFields = []) {
  const registryMap = new Map();
  const categoryBuckets = new Map();

  (Array.isArray(mergeFields) ? mergeFields : []).forEach(field => {
    if (!field || !field.field_key) return;
    registryMap.set(field.field_key, field);
    const groupKey = CATEGORY_TO_GROUP_KEY[field.category];
    if (!groupKey) return;
    if (field.active === 0 || field.active === false) return;
    if (field.show_in_jobsheet === 0 || field.show_in_jobsheet === false) return;
    if (!categoryBuckets.has(groupKey)) {
      categoryBuckets.set(groupKey, []);
    }
    categoryBuckets.get(groupKey).push(field);
  });

  const groups = [];

  Object.entries(GROUP_CONFIG).forEach(([groupKey, config]) => {
    const fields = [];
    const used = new Set();
    const bucket = categoryBuckets.get(groupKey) || [];
    const bucketMap = new Map(bucket.map(field => [field.field_key, field]));
    const hasRegistryData = registryMap.size > 0;

    const addField = (fieldKey, { force } = {}) => {
      if (!fieldKey || used.has(fieldKey)) return;
      const meta = FIELD_META[fieldKey] || {};
      const candidate = bucketMap.get(fieldKey) || registryMap.get(fieldKey);
      const showFromRegistry = Boolean(candidate && candidate.show_in_jobsheet !== false && candidate.active !== false);
      const fieldExistsInRegistry = registryMap.has(fieldKey);
      const shouldInclude =
        force ||
        meta.always ||
        showFromRegistry ||
        (!fieldExistsInRegistry && meta.fallback) ||
        (!hasRegistryData && meta.fallback);

      if (!shouldInclude) return;

      const fieldConfig = buildFieldConfig(fieldKey, candidate);
      if (!fieldConfig) return;

      if (!fieldConfig.hint && candidate?.description) {
        fieldConfig.hint = candidate.description;
      }

      fields.push(fieldConfig);
      used.add(fieldKey);
    };

    if (config.staticOnly) {
      (config.fields || []).forEach(key => addField(key, { force: true }));
    } else {
      (config.prepend || []).forEach(key => addField(key, { force: true }));
      (config.order || []).forEach(key => addField(key));
      bucket
        .filter(field => !used.has(field.field_key))
        .sort((a, b) => {
          const labelA = a.label || startCaseKey(a.field_key);
          const labelB = b.label || startCaseKey(b.field_key);
          return labelA.localeCompare(labelB);
        })
        .forEach(field => addField(field.field_key));
      (config.append || []).forEach(key => addField(key, { force: true }));
    }

    if (fields.length > 0) {
      groups.push({
        key: groupKey,
        title: config.title,
        description: config.description,
        defaultOpen: Boolean(config.defaultOpen),
        fields,
        icon: GROUP_ICON_MAP[groupKey] || '📄'
      });
    }
  });

  return groups;
}

const FALLBACK_JOBSHEET_GROUPS = buildJobsheetGroups([]);

function formatDateInput(value) {
  if (!value) return '';
  const parsed = parseDateValue(value);
  if (!parsed) return '';
  return parsed.toISOString().slice(0, 10);
}

function formatDateDisplay(value) {
  const parsed = parseDateValue(value);
  if (!parsed) return 'Date tbc';
  return parsed.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'long',
    year: 'numeric'
  });
}

function formatTimestampDisplay(value) {
  if (!value) return '—';
  const parsed = parseDateValue(value);
  if (!parsed) return typeof value === 'string' ? value : '—';
  const datePart = parsed.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });
  const timePart = parsed.toLocaleTimeString('en-GB', {
    hour: '2-digit',
    minute: '2-digit'
  });
  return `${datePart} ${timePart}`;
}

function formatFileSize(value) {
  const bytes = Number(value);
  if (!Number.isFinite(bytes) || bytes < 0) return '—';
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  const precision = size >= 10 || unitIndex === 0 ? 0 : 1;
  return `${size.toFixed(precision)} ${units[unitIndex]}`;
}

function normalizeLookupString(value) {
  return (value || '')
    .toString()
    .trim()
    .toLowerCase();
}

function normalizeDateKey(value) {
  const iso = formatDateInput(value);
  return iso || '';
}

function slugifyForMatch(value) {
  return (value || '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .trim();
}

function matchesDocumentToJobsheet(doc, jobsheetState) {
  if (!doc || !jobsheetState) return false;

  const jobClient = normalizeLookupString(jobsheetState.client_name);
  const docClient = normalizeLookupString(
    doc.client_name
    || doc.display_client_name
    || doc.joined_client_name
  );

  const jobEventDate = normalizeDateKey(jobsheetState.event_date);
  const docEventDate = normalizeDateKey(
    doc.event_date
    || doc.display_event_date
    || doc.joined_event_date
  );

  const jobEventName = normalizeLookupString(jobsheetState.event_type);
  const docEventName = normalizeLookupString(
    doc.event_name
    || doc.display_event_name
    || doc.joined_event_name
  );

  const clientMatches = Boolean(jobClient && docClient && docClient === jobClient);
  const eventDateMatches = Boolean(jobEventDate && docEventDate && docEventDate === jobEventDate);
  const eventNameMatches = Boolean(jobEventName && docEventName && docEventName === jobEventName);

  if (clientMatches && (eventDateMatches || !jobEventDate || !docEventDate)) return true;
  if (eventDateMatches && (clientMatches || !jobClient || !docClient)) return true;
  if (clientMatches && eventNameMatches) return true;

  const jobSlug = slugifyForMatch(jobsheetState.client_name || jobsheetState.event_type);
  if (jobSlug && typeof doc.file_path === 'string' && doc.file_path.toLowerCase().includes(jobSlug)) {
    return true;
  }

  return false;
}

function getGroupIcon(groupKey) {
  return GROUP_ICON_MAP[groupKey] || '📄';
}

function getWorkspaceIcon(sectionKey) {
  return WORKSPACE_ICON_MAP[sectionKey] || '🗂️';
}

function fuzzyScore(query, text) {
  if (!query) return 0;
  if (!text) return null;
  const haystack = text.toString().toLowerCase();
  const needle = query.toLowerCase();
  let score = 0;
  let lastIndex = -1;

  for (let i = 0; i < needle.length; i += 1) {
    const char = needle[i];
    const matchIndex = haystack.indexOf(char, lastIndex + 1);
    if (matchIndex === -1) return null;
    if (lastIndex === -1) {
      score += matchIndex;
    } else {
      score += Math.max(0, matchIndex - lastIndex - 1);
    }
    if (matchIndex === i) {
      score -= 0.5;
    }
    lastIndex = matchIndex;
  }

  if (haystack.includes(needle)) {
    score -= 2;
  }

  return score;
}

function buildJobsheetHaystacks(sheet) {
  const items = [
    sheet.client_name,
    sheet.client_email,
    sheet.event_type,
    sheet.venue_name,
    sheet.venue_town,
    sheet.notes,
    sheet.service_types,
    sheet.status,
    sheet.jobsheet_id != null ? `#${sheet.jobsheet_id}` : '',
    sheet.pricing_service_id
  ].filter(Boolean);

  const combined = items.join(' ');
  return [...items, combined];
}

function getComparableValueForSort(sheet, field) {
  switch (field) {
    case 'event_date':
      return sheet.event_date ? new Date(sheet.event_date).valueOf() : 0;
    case 'ahmen_fee': {
      const total = Number(sheet.pricing_total);
      if (Number.isFinite(total) && total > 0) return total;
      const singerFee = Number(sheet.ahmen_fee) || 0;
      const productionFee = Number(sheet.production_fees) || 0;
      return singerFee + productionFee;
    }
    case 'status':
      return STATUS_ORDER[sheet.status] ?? STATUS_OPTIONS.length;
    case 'client_name':
    case 'event_type':
      return (sheet[field] || '').toString().toLowerCase();
    case 'venue_name':
      return (sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '').toString().toLowerCase();
    default:
      return sheet[field];
  }
}

function IconButton({ label, onClick, disabled, className = '', children, size = 'md' }) {
  const handleClick = useCallback((event) => {
    event.stopPropagation();
    onClick?.(event);
  }, [onClick]);

  const wantsCustomColors = /\b(border-|text-|hover:bg-)\w/.test(className || '');
  const colorClasses = wantsCustomColors ? '' : 'border-slate-300 text-slate-600 hover:bg-slate-100';
  const sizeClasses = size === 'sm' ? 'h-7 w-7' : (size === 'lg' ? 'h-10 w-10' : 'h-9 w-9');

  return (
    <button
      type="button"
      onClick={handleClick}
      disabled={disabled}
      className={`inline-flex ${sizeClasses} items-center justify-center rounded border transition disabled:cursor-not-allowed disabled:opacity-60 ${colorClasses} ${className}`}
      aria-label={label}
      title={label}
    >
      {children}
    </button>
  );
}

function TreeActionButton({ title, onClick, disabled, children }) {
  const handleClick = (event) => {
    event.stopPropagation();
    if (!disabled) {
      onClick?.(event);
    }
  };

  return (
    <button
      type="button"
      onClick={handleClick}
      disabled={disabled}
      className="rounded border border-transparent p-1 text-indigo-500 transition hover:bg-indigo-100 hover:text-indigo-700 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-indigo-500 disabled:cursor-not-allowed disabled:opacity-40"
      title={title}
      aria-label={title}
    >
      {children}
    </button>
  );
}

// InvoiceNumberingCard removed per filename-driven numbering model

function InvoiceLogPanel({ business, onOpenFile, onRevealFile, onDeleteDocument }) {
  const businessId = business?.id;
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [list, setList] = useState([]);
  const [filter, setFilter] = useState('all'); // all | unpaid | overdue | duesoon | paid
  const [search, setSearch] = useState('');
  const [importing, setImporting] = useState(false);
  const [selectedInvoiceId, setSelectedInvoiceId] = useState(null);
  // More inline edit state
  const [editingAmountId, setEditingAmountId] = useState(null);
  const [amountDraft, setAmountDraft] = useState('');
  const [savingAmountId, setSavingAmountId] = useState(null);
  const [editingClientId, setEditingClientId] = useState(null);
  const [clientDraft, setClientDraft] = useState('');
  const [eventNameDraft, setEventNameDraft] = useState('');
  const [eventDateDraft, setEventDateDraft] = useState('');
  const [savingClientId, setSavingClientId] = useState(null);
  const [editingStatusId, setEditingStatusId] = useState(null);
  const [statusDraft, setStatusDraft] = useState('Issued');
  const [savingStatusId, setSavingStatusId] = useState(null);
  // Rebuild preview state
  const [rebuildOpen, setRebuildOpen] = useState(false);
  const [rebuildDoc, setRebuildDoc] = useState(null);
  const [rebuildPreview, setRebuildPreview] = useState(null);
  const [rebuildLoading, setRebuildLoading] = useState(false);
  // Relink to jobsheet state
  const [relinkOpen, setRelinkOpen] = useState(false);
  const [relinkDoc, setRelinkDoc] = useState(null);
  const [relinkLoading, setRelinkLoading] = useState(false);
  const [relinkJobsheets, setRelinkJobsheets] = useState([]);
  const [relinkSearch, setRelinkSearch] = useState('');
  const [relinkSelectedId, setRelinkSelectedId] = useState(null);
  const [relinkPreview, setRelinkPreview] = useState(null);
  // Inline edit state
  const [editingNumberId, setEditingNumberId] = useState(null);
  const [numberDraft, setNumberDraft] = useState('');
  const [savingNumberId, setSavingNumberId] = useState(null);
  const [editingDueId, setEditingDueId] = useState(null);
  const [dueDraft, setDueDraft] = useState('');
  const [savingDueId, setSavingDueId] = useState(null);
  const [editingRemId, setEditingRemId] = useState(null);
  const [remDraft, setRemDraft] = useState('');
  const [savingRemId, setSavingRemId] = useState(null);
  const [sortKey, setSortKey] = useState('due'); // number | client | amount | due | reminder | status
  const [sortDir, setSortDir] = useState('desc'); // asc | desc

  const toggleSort = useCallback((key) => {
    setSortKey(prevKey => {
      if (prevKey !== key) {
        // default directions per column
        setSortDir((key === 'client' || key === 'status') ? 'asc' : 'desc');
        return key;
      }
      setSortDir(prev => (prev === 'asc' ? 'desc' : 'asc'));
      return key;
    });
  }, []);

  const renderSortIndicator = useCallback((key) => {
    if (sortKey !== key) return null;
    return (
      <span className="ml-1 inline-block text-slate-400">{sortDir === 'asc' ? '▲' : '▼'}</span>
    );
  }, [sortKey, sortDir]);

  

  const refresh = useCallback(async () => {
    if (!businessId) return;
    try {
      setLoading(true);
      setError('');
      // First, index any PDFs named with (INV-###) into invoice rows
      try { await window.api?.indexInvoicesFromFilenames?.({ businessId }); } catch (_err) {}
      // Reconcile DB with filesystem to avoid ghost files before loading
      try { await window.api?.cleanOrphanDocuments?.({ businessId }); } catch (_err) {}
      const allDocs = await window.api?.getDocuments?.({ businessId });
      const docs = Array.isArray(allDocs) ? allDocs : [];

      const isInvoiceLike = (doc) => {
        if (!doc) return false;
        const type = String(doc.doc_type || '').toLowerCase();
        if (type === 'invoice') return true;
        if (type !== 'pdf_export') return false;
        const fp = String(doc.file_path || '').toLowerCase();
        const name = String(doc.file_name || doc.definition_label || doc.label || '').toLowerCase();
        const hay = fp || name;
        return hay.includes('invoice') && (hay.includes('deposit') || hay.includes('balance'));
      };

      // Filter by invoice-like documents
      const invoiceLike = docs.filter(isInvoiceLike);

      // Deduplicate by file_path, prefer real invoice rows over pdf_export
      const byPath = new Map();
      invoiceLike.forEach(d => {
        const key = String(d.file_path || '').trim();
        const existing = key ? byPath.get(key) : null;
        if (!existing) { byPath.set(key, d); return; }
        const existingIsInvoice = String(existing.doc_type || '').toLowerCase() === 'invoice';
        const currentIsInvoice = String(d.doc_type || '').toLowerCase() === 'invoice';
        if (!existingIsInvoice && currentIsInvoice) byPath.set(key, d);
      });
      const merged = Array.from(byPath.values());

      // Only show items whose files exist
      let filtered = merged;
      try {
        filtered = await window.api?.filterDocumentsByExistingFiles?.(merged, { includeMissing: false });
      } catch (_err) {
        filtered = merged.filter(d => d && d.file_path);
      }
      // Enrich missing totals/dates from jobsheets when available
      const variantOf = (doc) => {
        const a = String(doc.definition_invoice_variant || doc.invoice_variant || '').toLowerCase();
        if (a === 'deposit' || a === 'balance') return a;
        const fp = String(doc.file_path || '').toLowerCase();
        const label = String(doc.display_label || doc.label || '').toLowerCase();
        const hay = `${fp} ${label}`;
        if (hay.includes('deposit')) return 'deposit';
        if (hay.includes('balance')) return 'balance';
        return '';
      };

      const ids = Array.from(new Set((filtered || []).map(d => d?.jobsheet_id).filter(id => id != null).map(Number)));
      const jsMap = new Map();
      for (const jid of ids) {
        try {
          // eslint-disable-next-line no-await-in-loop
          const js = await window.api?.getAhmenJobsheet?.(jid);
          if (js) jsMap.set(jid, js);
        } catch (_) {}
      }

      const enriched = (filtered || []).map(d => {
        let amount = d?.total_amount ?? d?.balance_due ?? null;
        let due_date = d?.due_date ?? null;
        let reminder_date = d?.reminder_date ?? null;
        if ((!Number.isFinite(Number(amount)) || Number(amount) === 0) || (!due_date && !reminder_date)) {
          const js = d?.jobsheet_id != null ? jsMap.get(Number(d.jobsheet_id)) : null;
          const v = variantOf(d);
          if (js && v) {
            if (v === 'deposit') {
              amount = js.deposit_amount != null ? Number(js.deposit_amount) : amount;
              // Deposit: no reminder, due on contract signing (leave blank)
              reminder_date = null;
            } else if (v === 'balance') {
              amount = js.balance_amount != null ? Number(js.balance_amount) : amount;
              due_date = js.balance_due_date || due_date;
              reminder_date = js.balance_reminder_date != null ? js.balance_reminder_date : reminder_date;
            }
          }
        }
        return { ...d, total_amount: amount, due_date, reminder_date };
      });

      setList(enriched || []);
    } catch (err) {
      console.error('Failed to load invoices', err);
      setError(err?.message || 'Unable to load invoices');
    } finally {
      setLoading(false);
    }
  }, [businessId]);

  useEffect(() => { refresh(); }, [refresh]);

  // Keep selection valid when list changes
  useEffect(() => {
    if (selectedInvoiceId == null) return;
    const exists = (Array.isArray(list) ? list : []).some(d => d && d.document_id === selectedInvoiceId);
    if (!exists) setSelectedInvoiceId(null);
  }, [list, selectedInvoiceId]);

  const handleImportHistoric = useCallback(async () => {
    if (!businessId || !window.api || typeof window.api.indexInvoicesFromFilenames !== 'function') return;
    try {
      setImporting(true);
      const result = await window.api.indexInvoicesFromFilenames({ businessId });
      await refresh();
      const count = result && typeof result.imported === 'number' ? result.imported : 0;
      window.alert(`Imported ${count} invoice${count === 1 ? '' : 's'} from filenames`);
    } catch (err) {
      console.error('Historic import failed', err);
      setError(err?.message || 'Unable to import historic invoices');
    } finally {
      setImporting(false);
    }
  }, [businessId, refresh]);

  // Auto-refresh on document change events and jobsheet document updates
  useEffect(() => {
    if (!businessId || !window.api) return () => {};
    // Ensure a watcher is running for this business (idempotent)
    window.api.watchDocuments?.({ businessId }).catch(() => {});
    const unsubDocs = window.api.onDocumentsChange?.((payload) => {
      if (!payload || payload.businessId !== businessId) return;
      refresh();
    });
    const unsubJobsheet = window.api.onJobsheetChange?.((payload) => {
      if (!payload || payload.businessId !== businessId) return;
      if (payload.type === 'documents-updated') refresh();
    });
    return () => { unsubDocs?.(); unsubJobsheet?.(); };
  }, [businessId, refresh]);

  const toggleLock = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    try {
      await window.api?.setDocumentLock?.(doc.document_id, !doc.is_locked);
      refresh();
      try {
        window.api?.notifyJobsheetChange?.({
          type: 'documents-updated',
          businessId,
          jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
          documentId: doc.document_id
        });
      } catch (_) {}
    } catch (err) {
      console.error('Failed to toggle lock', err);
      setError(err?.message || 'Unable to toggle lock');
    }
  }, [refresh]);

  const handleMarkPaidToggle = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    try {
      const isPaid = String(doc.status || '').toLowerCase() === 'paid' || !!doc.paid_at;
      await window.api?.updateDocumentStatus?.(doc.document_id, isPaid ? { status: 'Issued', paid_at: null } : { status: 'Paid', paid_at: new Date().toISOString() });
      refresh();
      try {
        window.api?.notifyJobsheetChange?.({
          type: 'documents-updated',
          businessId,
          jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
          documentId: doc.document_id
        });
      } catch (_) {}
    } catch (err) {
      console.error('Failed to update payment status', err);
      setError(err?.message || 'Unable to update payment status');
    }
  }, [refresh]);

  const handleSetNumber = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const isInvoice = String(doc.doc_type || '').toLowerCase() === 'invoice';
    const doUnlock = async () => {
      if (!doc.is_locked) return true;
      const proceed = window.confirm('This record is locked. Unlock to continue?');
      if (!proceed) return false;
      try { await window.api?.setDocumentLock?.(doc.document_id, false); } catch (_) {}
      return true;
    };
    if (!(await doUnlock())) return;

    let requested = window.prompt('Set invoice number (leave blank for next)', isInvoice && doc.number != null ? String(doc.number) : '');
    if (requested === '') requested = null;

    try {
      if (requested == null) return; // cancelled
      const digits = String(requested).replace(/[^0-9]/g, '');
      if (!digits) return;
      const val = Number(digits);
      if (!Number.isInteger(val) || val < 1) return;

      if (isInvoice) {
        await window.api?.setDocumentNumber?.(doc.document_id, val);
      } else {
        await window.api?.promotePdfToInvoice?.(doc.document_id, { number: val });
      }
      refresh();
      try {
        window.api?.notifyJobsheetChange?.({
          type: 'documents-updated',
          businessId,
          jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
          documentId: doc.document_id
        });
      } catch (_) {}
    } catch (err) {
      console.error('Failed to apply invoice number', err);
      setError(err?.message || 'Unable to apply invoice number');
    }
  }, [refresh]);

  // Helpers for inline editing
  const ensureUnlocked = useCallback(async (doc) => {
    if (!doc?.is_locked) return true;
    const proceed = window.confirm('This record is locked. Unlock to edit?');
    if (!proceed) return false;
    try { await window.api?.setDocumentLock?.(doc.document_id, false); return true; } catch (_) { return false; }
  }, []);

  const beginEditNumber = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingNumberId(doc.document_id);
    setNumberDraft(doc.number != null ? String(doc.number) : '');
  }, [ensureUnlocked]);

  const commitEditNumber = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const raw = String(numberDraft || '').trim();
    const digits = raw.replace(/[^0-9]/g, '');
    const isInvoice = String(doc.doc_type || '').toLowerCase() === 'invoice';
    if (!digits) { return; }
    const val = Number(digits);
    if (!Number.isInteger(val) || val < 1) { return; }
    try {
      setSavingNumberId(doc.document_id);
      if (isInvoice) {
        await window.api?.setDocumentNumber?.(doc.document_id, val);
      } else {
        await window.api?.promotePdfToInvoice?.(doc.document_id, { number: val });
      }
      await refresh();
      setEditingNumberId(null);
    } catch (err) {
      // Non-blocking error; keep editing so the user can retry
    } finally { setSavingNumberId(null); }
  }, [numberDraft, refresh]);

  const cancelEditNumber = useCallback(() => { setEditingNumberId(null); setNumberDraft(''); }, []);

  const toInputDate = useCallback((v) => {
    if (!v) return '';
    const d = new Date(v);
    return Number.isNaN(d.valueOf()) ? '' : d.toISOString().slice(0,10);
  }, []);

  const beginEditDue = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingDueId(doc.document_id);
    setDueDraft(toInputDate(doc.due_date));
  }, [ensureUnlocked, toInputDate]);
  const beginEditRem = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingRemId(doc.document_id);
    setRemDraft(toInputDate(doc.reminder_date));
  }, [ensureUnlocked, toInputDate]);

  const commitEditDue = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const payload = { due_date: dueDraft ? dueDraft : null };
    try { setSavingDueId(doc.document_id); await window.api?.updateDocumentStatus?.(doc.document_id, payload); await refresh(); setEditingDueId(null); } catch (err) { window.alert(err?.message || 'Unable to set due date'); } finally { setSavingDueId(null); }
  }, [dueDraft, refresh]);
  const cancelEditDue = useCallback(() => { setEditingDueId(null); setDueDraft(''); }, []);

  const commitEditRem = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const payload = { reminder_date: remDraft ? remDraft : null };
    try { setSavingRemId(doc.document_id); await window.api?.updateDocumentStatus?.(doc.document_id, payload); await refresh(); setEditingRemId(null); } catch (err) { window.alert(err?.message || 'Unable to set reminder date'); } finally { setSavingRemId(null); }
  }, [remDraft, refresh]);
  const cancelEditRem = useCallback(() => { setEditingRemId(null); setRemDraft(''); }, []);

  const beginEditAmount = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingAmountId(doc.document_id);
    const base = (doc.total_amount ?? doc.balance_due);
    setAmountDraft(base != null ? String(base) : '');
  }, [ensureUnlocked]);

  const commitEditAmount = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const raw = String(amountDraft || '').trim();
    if (!raw) { setEditingAmountId(null); return; }
    // Allow decimals but normalize to 2dp
    const val = Number(raw);
    if (!Number.isFinite(val) || val < 0) { setEditingAmountId(null); return; }
    try {
      setSavingAmountId(doc.document_id);
      const rounded = Math.round(val * 100) / 100;
      await window.api?.updateDocumentStatus?.(doc.document_id, { total_amount: rounded, balance_due: rounded });
      await refresh();
      setEditingAmountId(null);
    } catch (err) {
      // keep silent
    } finally { setSavingAmountId(null); }
  }, [amountDraft, refresh]);
  const cancelEditAmount = useCallback(() => { setEditingAmountId(null); setAmountDraft(''); }, []);

  const beginEditClient = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingClientId(doc.document_id);
    setClientDraft(doc.client_name || doc.display_client_name || '');
    setEventNameDraft(doc.event_name || doc.display_event_name || '');
    setEventDateDraft(toInputDate(doc.display_event_date || doc.event_date || ''));
  }, [ensureUnlocked, toInputDate]);

  const commitEditClient = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const payload = {
      client_name: (clientDraft || '').trim(),
      event_name: (eventNameDraft || '').trim(),
      event_date: (eventDateDraft || '').trim() || null
    };
    try {
      setSavingClientId(doc.document_id);
      await window.api?.updateDocumentStatus?.(doc.document_id, payload);
      await refresh();
      setEditingClientId(null);
    } catch (err) {
      // ignore
    } finally { setSavingClientId(null); }
  }, [clientDraft, eventNameDraft, eventDateDraft, refresh]);
  const cancelEditClient = useCallback(() => { setEditingClientId(null); setClientDraft(''); setEventNameDraft(''); setEventDateDraft(''); }, []);

  const beginEditStatus = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    if (!(await ensureUnlocked(doc))) return;
    setEditingStatusId(doc.document_id);
    const curr = String(doc.status || '').toLowerCase() === 'paid' ? 'Paid' : 'Issued';
    setStatusDraft(curr);
  }, [ensureUnlocked]);

  const commitEditStatus = useCallback(async (doc) => {
    if (!doc?.document_id) return;
    const v = String(statusDraft || 'Issued');
    const payload = v === 'Paid' ? { status: 'Paid', paid_at: new Date().toISOString() } : { status: 'Issued', paid_at: null };
    try {
      setSavingStatusId(doc.document_id);
      await window.api?.updateDocumentStatus?.(doc.document_id, payload);
      await refresh();
      setEditingStatusId(null);
    } catch (err) {
      // ignore
    } finally { setSavingStatusId(null); }
  }, [statusDraft, refresh]);
  const cancelEditStatus = useCallback(() => { setEditingStatusId(null); setStatusDraft('Issued'); }, []);

  const handleOpen = useCallback(async (doc) => {
    const filePath = doc?.file_path || '';
    if (!filePath) { setError('PDF not available for this invoice'); return; }
    try {
      setError('');
      const res = await window.api?.openPath?.(filePath);
      if (res && res.ok === false) throw new Error(res.message || 'Unable to open file');
    } catch (err) {
      console.error('Open failed', err);
      setError(err?.message || 'Unable to open file');
    }
  }, []);

  const handleReveal = useCallback(async (doc) => {
    const filePath = doc?.file_path || '';
    if (!filePath) { setError('PDF not available for this invoice'); return; }
    try {
      setError('');
      const res = await window.api?.showItemInFolder?.(filePath);
      if (res && res.ok === false) throw new Error(res.message || 'Unable to reveal file');
    } catch (err) {
      console.error('Reveal failed', err);
      setError(err?.message || 'Unable to reveal file');
    }
  }, []);

  const quickLookSelected = useCallback(async () => {
    const id = selectedInvoiceId != null ? Number(selectedInvoiceId) : null;
    if (!id) return;
    const doc = computed.find(d => d && Number(d.document_id) === id);
    const filePath = doc?.file_path || '';
    if (!filePath) return;
    try {
      await window.api?.quickLookPath?.(filePath);
    } catch (_err) {
      try { await window.api?.openPath?.(filePath); } catch (_) {}
    }
  }, [selectedInvoiceId, computed]);

  const panelRef = useRef(null);

  // Persist sort
  const SORT_STORAGE_KEY = useMemo(() => `ui:${businessId || 0}:invoiceLogSort`, [businessId]);
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(SORT_STORAGE_KEY);
      if (raw) {
        const p = JSON.parse(raw);
        if (p && typeof p === 'object') {
          if (p.key) setSortKey(p.key);
          if (p.dir) setSortDir(p.dir === 'asc' ? 'asc' : 'desc');
        }
      }
    } catch (_) {}
    // load once per business
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [SORT_STORAGE_KEY]);
  useEffect(() => {
    try {
      const persist = window.localStorage.getItem('app:persistUiState') === 'true';
      if (!persist) return;
      window.localStorage.setItem(SORT_STORAGE_KEY, JSON.stringify({ key: sortKey, dir: sortDir }));
    } catch (_) {}
  }, [SORT_STORAGE_KEY, sortKey, sortDir]);

  const computed = useMemo(() => {
    const today = new Date();
    const seven = 7 * 24 * 60 * 60 * 1000;
    const rows = (Array.isArray(list) ? list : []).map(d => {
      const parseInv = () => {
        const fp = String(d.file_path || '');
        const base = fp ? fp.split(/[\\/]+/).pop() || '' : '';
        const m = base.match(/INV[-\s]?(\d+)/i);
        return m ? Number(m[1]) : null;
      };
      const numberParsed = d?.number != null ? Number(d.number) : parseInv();
      const due = d?.due_date ? new Date(d.due_date) : null;
      const paid = String(d.status || '').toLowerCase() === 'paid' || !!d.paid_at;
      let derived = paid ? 'Paid' : 'Issued';
      if (!paid && due instanceof Date && !Number.isNaN(due.valueOf())) {
        if (due.valueOf() < today.valueOf()) derived = 'Overdue';
        else if (due.valueOf() - today.valueOf() <= seven) derived = 'Due soon';
      }
      return { ...d, derived_status: derived, number: d.number != null ? d.number : numberParsed };
    });

    const s = (search || '').trim().toLowerCase();
    let filtered = rows;
    if (filter !== 'all') {
      filtered = rows.filter(r => {
        const st = (r.derived_status || '').toLowerCase();
        if (filter === 'paid') return st === 'paid';
        if (filter === 'unpaid') return st !== 'paid';
        if (filter === 'overdue') return st === 'overdue';
        if (filter === 'duesoon') return st === 'due soon';
        return true;
      });
    }
    if (s) {
      filtered = filtered.filter(r => {
        const hay = [
          r.number != null ? String(r.number) : '',
          r.display_client_name || r.client_name || '',
          r.display_event_name || r.event_name || ''
        ].join(' ').toLowerCase();
        return hay.includes(s);
      });
    }
    // Sorting
    const cmp = (a, b) => {
      const dirMul = sortDir === 'asc' ? 1 : -1;
      const safeStr = (v) => (v || '').toString().toLowerCase();
      const asDate = (v) => {
        if (!v) return null;
        const d = new Date(v);
        return Number.isNaN(d.valueOf()) ? null : d.valueOf();
      };
      if (sortKey === 'number') {
        const av = Number(a.number);
        const bv = Number(b.number);
        const aValid = Number.isFinite(av);
        const bValid = Number.isFinite(bv);
        if (aValid && bValid) return (av - bv) * dirMul;
        if (aValid) return -1 * dirMul; // numbers first
        if (bValid) return 1 * dirMul;
        return 0;
      }
      if (sortKey === 'client') {
        const av = safeStr(a.display_client_name || a.client_name);
        const bv = safeStr(b.display_client_name || b.client_name);
        if (av !== bv) return av.localeCompare(bv) * dirMul;
        // tiebreaker: event date
        const ad = asDate(a.display_event_date || a.event_date);
        const bd = asDate(b.display_event_date || b.event_date);
        if (ad != null && bd != null) return (ad - bd) * dirMul;
        return 0;
      }
      if (sortKey === 'amount') {
        const av = Number(a.total_amount ?? a.balance_due);
        const bv = Number(b.total_amount ?? b.balance_due);
        return ((Number.isFinite(av) ? av : -Infinity) - (Number.isFinite(bv) ? bv : -Infinity)) * dirMul;
      }
      if (sortKey === 'due') {
        const av = asDate(a.due_date);
        const bv = asDate(b.due_date);
        if (av != null && bv != null) return (av - bv) * dirMul;
        if (av != null) return -1 * dirMul;
        if (bv != null) return 1 * dirMul;
        return 0;
      }
      if (sortKey === 'reminder') {
        const av = asDate(a.reminder_date);
        const bv = asDate(b.reminder_date);
        if (av != null && bv != null) return (av - bv) * dirMul;
        if (av != null) return -1 * dirMul;
        if (bv != null) return 1 * dirMul;
        return 0;
      }
      if (sortKey === 'status') {
        const order = { paid: 3, 'due soon': 2, overdue: 1, issued: 0 };
        const av = order[(a.derived_status || '').toLowerCase()] ?? -1;
        const bv = order[(b.derived_status || '').toLowerCase()] ?? -1;
        if (av !== bv) return (av - bv) * dirMul;
        return 0;
      }
      return 0;
    };
    return filtered.slice().sort(cmp);
  }, [list, filter, search, sortKey, sortDir]);

  // Bulk lock all visible (based on current filter/search)
  const canLockAll = useMemo(() => {
    const rows = Array.isArray(computed) ? computed : [];
    return rows.some(d => d?.document_id != null && !d?.is_locked);
  }, [computed]);

  const handleLockAll = useCallback(async () => {
    try {
      const rows = Array.isArray(computed) ? computed : [];
      const ids = rows.filter(d => d?.document_id != null && !d?.is_locked).map(d => d.document_id);
      if (!ids.length) return;
      const ok = window.confirm(`Lock ${ids.length} invoice${ids.length === 1 ? '' : 's'}?`);
      if (!ok) return;
      for (const id of ids) {
        // eslint-disable-next-line no-await-in-loop
        await window.api?.setDocumentLock?.(id, true);
      }
      await refresh();
    } catch (err) {
      window.alert(err?.message || 'Unable to lock invoices');
    }
  }, [computed, refresh]);

  const formatDate = (val) => {
    if (!val) return '';
    const d = new Date(val);
    if (Number.isNaN(d.valueOf())) return '';
    return new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).format(d);
  };

  const toCurrency = (val) => {
    const n = Number(val);
    if (!Number.isFinite(n)) return '';
    try { return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(n); } catch (_err) { return `£${n.toFixed(2)}`; }
  };

  // Narrower INV column, wider Client/Event column
  const GRID_COLS = '0.8fr 3.4fr 1fr 0.9fr 1fr 0.9fr 64px';
  const GRID_TEMPLATE_WITH_SEP = '0.8fr 1px 3.4fr 1px 1fr 1px 0.9fr 1px 1fr 1px 0.9fr 1px 64px';
  const CELL_BORDER = '#94a3b8'; // clearer divider (slate-400)

  const statusPill = (value) => {
    const v = (value || '').toLowerCase();
    const palette = {
      paid:   { bg: '#ecfdf5', border: '#86efac', color: '#166534' }, // green-50/300/800-ish
      overdue:{ bg: '#fef2f2', border: '#fca5a5', color: '#991b1b' }, // red-50/300/800-ish
      'due soon': { bg: '#fffbeb', border: '#fcd34d', color: '#92400e' }, // amber-50/300/800-ish
      default:{ bg: '#f8fafc', border: '#cbd5e1', color: '#334155' } // slate-50/300/700-ish
    };
    const p = palette[v] || palette.default;
    return (
      <span
        className="inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium border"
        style={{ backgroundColor: p.bg, borderColor: p.border, color: p.color }}
      >
        {value || '—'}
      </span>
    );
  };

  const [menuOpenId, setMenuOpenId] = useState(null);
  const closeMenus = useCallback(() => setMenuOpenId(null), []);
  useEffect(() => {
    const onDoc = (e) => {
      if (!menuOpenId) return;
      try {
        const target = e.target;
        if (target && typeof target.closest === 'function') {
          const container = target.closest('[data-kebab-for]');
          if (container) {
            const idAttr = container.getAttribute('data-kebab-for');
            const idNum = idAttr != null ? Number(idAttr) : null;
            if (idNum != null && idNum === menuOpenId) {
              // Click was inside the open menu/button container; do not auto-close
              return;
            }
          }
        }
      } catch (_) {}
      closeMenus();
    };
    document.addEventListener('scroll', onDoc, true);
    document.addEventListener('mousedown', onDoc);
    return () => { document.removeEventListener('scroll', onDoc, true); document.removeEventListener('mousedown', onDoc); };
  }, [menuOpenId, closeMenus]);

  // Close kebab on Escape
  useEffect(() => {
    const onKey = (e) => {
      if (!menuOpenId) return;
      if (e.key === 'Escape' || e.key === 'Esc') {
        e.stopPropagation();
        closeMenus();
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [menuOpenId, closeMenus]);

  return (<>
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-lg font-semibold text-slate-700">Invoice Log</h2>
          <p className="text-sm text-slate-500">Manage issued invoices and reminders.</p>
        </div>
          <div className="flex items-center gap-2 text-sm">
            <input
              type="text"
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search by #, client, or event…"
              className="w-64 rounded border border-slate-300 px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
            />
            <select value={filter} onChange={e => setFilter(e.target.value)} className="rounded border border-slate-300 px-2 py-1 text-sm">
              <option value="all">All</option>
              <option value="unpaid">Unpaid</option>
              <option value="overdue">Overdue</option>
              <option value="duesoon">Due soon</option>
              <option value="paid">Paid</option>
            </select>
            <button type="button" onClick={handleImportHistoric} disabled={importing} className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60">{importing ? 'Importing…' : 'Import from filenames'}</button>
            <button type="button" onClick={refresh} className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50">Refresh</button>
            <button
              type="button"
              onClick={handleLockAll}
              disabled={!canLockAll}
              className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
              title="Lock all visible invoices"
            >
              Lock all
            </button>
          </div>
      </div>

      {error ? <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">{error}</div> : null}
      {loading ? <div className="text-sm text-slate-500">Loading…</div> : null}

      <div
        className="rounded border border-slate-200 overflow-x-auto"
        tabIndex={0}
        ref={panelRef}
        onMouseDown={(e) => {
          // Only focus container if clicking outside form fields
          const t = e.target;
          const tag = t && t.tagName ? t.tagName.toLowerCase() : '';
          const isForm = tag === 'input' || tag === 'textarea' || tag === 'select' || (t && t.isContentEditable);
          if (isForm) return;
          try { panelRef.current?.focus({ preventScroll: true }); } catch (_) {}
        }}
        onKeyDown={(e) => {
          // Quick Look selected invoice with Spacebar (not while editing)
          if ((e.key === ' ' || e.code === 'Space') && !rebuildOpen) {
            const t = e.target;
            const tag = t && t.tagName ? t.tagName.toLowerCase() : '';
            const isForm = tag === 'input' || tag === 'textarea' || tag === 'select' || (t && t.isContentEditable);
            if (isForm) return; // let form fields handle space
            e.preventDefault();
            quickLookSelected();
          }
        }}
      >
        <div
          className="grid items-center px-3 py-3 text-xs font-semibold text-slate-600"
          style={{ gridTemplateColumns: GRID_COLS, textAlign: 'left' }}
        >
          <button type="button" onClick={() => toggleSort('number')} className="text-left hover:text-indigo-600">
            Invoice # {renderSortIndicator('number')}
          </button>
          <button type="button" onClick={() => toggleSort('client')} className="text-left hover:text-indigo-600">
            Client / Event {renderSortIndicator('client')}
          </button>
          <button type="button" onClick={() => toggleSort('amount')} className="text-left hover:text-indigo-600">
            Amount {renderSortIndicator('amount')}
          </button>
          <button type="button" onClick={() => toggleSort('due')} className="text-left hover:text-indigo-600">
            Due {renderSortIndicator('due')}
          </button>
          <button type="button" onClick={() => toggleSort('reminder')} className="text-left hover:text-indigo-600">
            Reminder {renderSortIndicator('reminder')}
          </button>
          <button type="button" onClick={() => toggleSort('status')} className="text-center hover:text-indigo-600">
            Status {renderSortIndicator('status')}
          </button>
          <div style={{ textAlign: 'center' }}>Actions</div>
        </div>
        <div className="divide-y divide-slate-200">
          {computed.map(doc => {
            const detectVariant = (d) => {
              const raw = String(d.definition_invoice_variant || d.invoice_variant || '').toLowerCase();
              if (raw === 'deposit' || raw === 'balance') return raw;
              const hay = `${String(d.file_path || '').toLowerCase()} ${String(d.display_label || d.label || '').toLowerCase()}`;
              if (hay.includes('deposit')) return 'deposit';
              if (hay.includes('balance')) return 'balance';
              return '';
            };
            const variant = detectVariant(doc);
            const variantLabel = variant === 'deposit' ? 'Deposit' : (variant === 'balance' ? 'Balance' : '');
            const numberLabel = doc.number != null ? `INV-${doc.number}` : '—';
            const title = numberLabel; // keep tooltip consistent with visible value
            const eventDateLabel = formatDate(doc.display_event_date || doc.event_date || '');
            const clientEventBase = [doc.display_client_name || doc.client_name || '', eventDateLabel || ''].filter(Boolean).join(' — ');
            const clientEvent = variantLabel ? `${clientEventBase} · ${variantLabel}` : clientEventBase;
            const status = doc.derived_status || doc.status || '';
            const locked = !!doc.is_locked;
            const statusKey = String(status || '').toLowerCase();
            const rowBg = (() => {
              if (statusKey === 'paid') return '#ecfdf5';
              if (statusKey === 'overdue') return '#fef2f2';
              if (statusKey === 'due soon') return '#fffbeb';
              return 'transparent';
            })();
            const isSelected = selectedInvoiceId != null && Number(selectedInvoiceId) === Number(doc.document_id);
            return (
              <div
                key={doc.document_id}
                className={`grid items-center px-3 py-3 text-sm relative ${isSelected ? 'ring-2 ring-indigo-300 rounded-md' : ''}`}
                style={{ gridTemplateColumns: GRID_COLS, backgroundColor: rowBg, textAlign: 'left' }}
                role="row"
                aria-selected={isSelected}
                onClick={() => setSelectedInvoiceId(doc.document_id)}
              >
                <div className="whitespace-normal break-words" title={title}>
                  <div className="inline-flex items-center gap-1 max-w-full">
                    {doc.is_locked ? (
                      <span className="text-slate-500" aria-label="Locked" title="Locked">🔒</span>
                    ) : null}
                    {editingNumberId === doc.document_id ? (
                      <input
                        type="number"
                        min={1}
                        step={1}
                        inputMode="numeric"
                        pattern="[0-9]*"
                        value={numberDraft}
                        disabled={savingNumberId === doc.document_id}
                        onChange={(e) => setNumberDraft(String(e.target.value || '').replace(/[^0-9]/g, ''))}
                        onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditNumber(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditNumber(); } }}
                        onBlur={() => commitEditNumber(doc)}
                        autoFocus
                        className="w-24 rounded border border-slate-300 px-2 py-0.5 text-sm"
                      />
                    ) : (
                      <button
                        type="button"
                        className={`rounded px-1 py-0.5 truncate ${doc.number == null ? 'text-indigo-600 hover:bg-indigo-50 underline' : 'text-slate-700 hover:bg-slate-50'}`}
                        onClick={() => beginEditNumber(doc)}
                        title={doc.number == null ? 'Set invoice number' : 'Edit invoice number'}
                      >
                        {numberLabel}
                      </button>
                    )}
                  </div>
                </div>
                <div className="truncate" title={clientEvent}>
                  {editingClientId === doc.document_id ? (
                    <div className="flex flex-col gap-1">
                      <input
                        type="text"
                        value={clientDraft}
                        disabled={savingClientId === doc.document_id}
                        onChange={(e) => setClientDraft(e.target.value)}
                        onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditClient(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditClient(); } }}
                        className="w-full rounded border border-slate-300 px-2 py-0.5 text-sm"
                        placeholder="Client name"
                        autoFocus
                      />
                      <div className="grid grid-cols-1 md:grid-cols-[1fr,auto] gap-1 items-center">
                        <input
                          type="text"
                          value={eventNameDraft}
                          disabled={savingClientId === doc.document_id}
                          onChange={(e) => setEventNameDraft(e.target.value)}
                          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditClient(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditClient(); } }}
                          className="w-full rounded border border-slate-300 px-2 py-0.5 text-sm"
                          placeholder="Event name"
                        />
                        <input
                          type="date"
                          value={eventDateDraft}
                          disabled={savingClientId === doc.document_id}
                          onChange={(e) => setEventDateDraft(e.target.value)}
                          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditClient(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditClient(); } }}
                          className="rounded border border-slate-300 px-2 py-0.5 text-sm"
                          placeholder="Event date"
                        />
                      </div>
                      <div className="flex items-center gap-2">
                        <button type="button" className="rounded border border-slate-300 px-2 py-0.5 text-xs text-slate-600 hover:bg-slate-50" onClick={() => commitEditClient(doc)}>Save</button>
                        <button type="button" className="rounded border border-slate-200 px-2 py-0.5 text-xs text-slate-500 hover:bg-slate-50" onClick={cancelEditClient}>Cancel</button>
                      </div>
                    </div>
                  ) : (
                    <button
                      type="button"
                      className="rounded px-1 py-0.5 text-slate-700 hover:bg-slate-50 w-full text-left"
                      onClick={() => beginEditClient(doc)}
                      title="Edit client/event"
                    >
                      {clientEvent}
                    </button>
                  )}
                </div>
                <div>
                  {editingAmountId === doc.document_id ? (
                    <input
                      type="number"
                      step="0.01"
                      value={amountDraft}
                      disabled={savingAmountId === doc.document_id}
                      onChange={(e) => setAmountDraft(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditAmount(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditAmount(); } }}
                      onBlur={() => commitEditAmount(doc)}
                      autoFocus
                      className="w-28 rounded border border-slate-300 px-2 py-0.5 text-sm text-right"
                    />
                  ) : (
                    <button type="button" className="rounded px-1 py-0.5 text-slate-700 hover:bg-slate-50 w-full text-right" onClick={() => beginEditAmount(doc)} title="Edit amount">
                      {toCurrency(doc.total_amount ?? doc.balance_due)}
                    </button>
                  )}
                </div>
                <div className="whitespace-nowrap">
                  {editingDueId === doc.document_id ? (
                    <input
                      type="date"
                      value={dueDraft}
                      disabled={savingDueId === doc.document_id}
                      onChange={(e) => setDueDraft(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditDue(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditDue(); } }}
                      onBlur={() => commitEditDue(doc)}
                      autoFocus
                      className="rounded border border-slate-300 px-2 py-0.5 text-sm"
                    />
                  ) : (
                    <button
                      type="button"
                      className="rounded px-1 py-0.5 text-slate-700 hover:bg-slate-50"
                      onClick={() => beginEditDue(doc)}
                      title="Edit due date"
                    >
                      {formatDate(doc.due_date) || '—'}
                    </button>
                  )}
                </div>
                <div className="whitespace-nowrap">
                  {editingRemId === doc.document_id ? (
                    <input
                      type="date"
                      value={remDraft}
                      disabled={savingRemId === doc.document_id}
                      onChange={(e) => setRemDraft(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commitEditRem(doc); } else if (e.key === 'Escape') { e.preventDefault(); cancelEditRem(); } }}
                      onBlur={() => commitEditRem(doc)}
                      autoFocus
                      className="rounded border border-slate-300 px-2 py-0.5 text-sm"
                    />
                  ) : (
                    <button
                      type="button"
                      className="rounded px-1 py-0.5 text-slate-700 hover:bg-slate-50"
                      onClick={() => beginEditRem(doc)}
                      title="Edit reminder date"
                    >
                      {formatDate(doc.reminder_date) || '—'}
                    </button>
                  )}
                </div>
                <div style={{ textAlign: 'center' }}>
                  {editingStatusId === doc.document_id ? (
                    <select
                      value={statusDraft}
                      disabled={savingStatusId === doc.document_id}
                      onChange={(e) => setStatusDraft(e.target.value)}
                      onBlur={() => commitEditStatus(doc)}
                      className="rounded border border-slate-300 px-2 py-1 text-xs"
                      autoFocus
                    >
                      <option value="Issued">Issued</option>
                      <option value="Paid">Paid</option>
                    </select>
                  ) : (
                    <button type="button" className="rounded px-2 py-0.5 text-slate-700 hover:bg-slate-50" onClick={() => beginEditStatus(doc)} title="Edit status">
                      {statusPill(status)}
                    </button>
                  )}
                </div>
                <div className="flex items-center justify-center gap-2 relative" data-kebab-for={doc.document_id}>
                  <IconButton size="md" label="Actions" className="bg-white" onClick={(e) => { e.stopPropagation(); setMenuOpenId(prev => prev === doc.document_id ? null : doc.document_id); }}>
                    <span aria-hidden>⋮</span>
                  </IconButton>
                  {menuOpenId === doc.document_id ? (
                    <div className="absolute right-0 top-full z-20 mt-2 w-48 rounded border border-slate-200 bg-white p-1 shadow-lg">
                      <button type="button" onClick={() => { toggleLock(doc); closeMenus(); }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">
                        {locked ? 'Unlock' : 'Lock'}
                      </button>
                      <button type="button" onClick={() => { handleSetNumber(doc); closeMenus(); }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">Set number…</button>
                      <button type="button" onClick={async () => {
                        closeMenus();
                        setRebuildDoc(doc);
                        setRebuildLoading(true);
                        setRebuildPreview(null);
                        setRebuildOpen(true);
                        try {
                          const res = await window.api?.rebuildInvoiceFromFilename?.({ businessId, documentId: doc.document_id, preview: true });
                          setRebuildPreview(res || null);
                        } catch (err) {
                          setRebuildPreview({ ok: false, message: err?.message || 'Unable to preview rebuild' });
                        } finally {
                          setRebuildLoading(false);
                        }
                      }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">Rebuild from filename…</button>
                      <button type="button" onClick={async () => {
                        closeMenus();
                        setRelinkDoc(doc);
                        setRelinkOpen(true);
                        setRelinkLoading(true);
                        setRelinkJobsheets([]);
                        setRelinkSelectedId(null);
                        setRelinkPreview(null);
                        try {
                          const list = await window.api?.getAhmenJobsheets?.({ businessId, includeArchived: true });
                          setRelinkJobsheets(Array.isArray(list) ? list : []);
                        } catch (_) {}
                        setRelinkLoading(false);
                      }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">Relink to jobsheet…</button>
                      <button type="button" onClick={async () => {
                        const current = doc?.due_date ? String(doc.due_date).slice(0,10) : '';
                        const next = window.prompt('Set due date (YYYY-MM-DD)', current);
                        if (next == null) { closeMenus(); return; }
                        try { await window.api?.updateDocumentStatus?.(doc.document_id, { due_date: next }); refresh(); } catch (err) { window.alert(err?.message || 'Unable to set due date'); }
                        closeMenus();
                      }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">Set due date…</button>
                      <button type="button" onClick={async () => {
                        const current = doc?.reminder_date ? String(doc.reminder_date).slice(0,10) : '';
                        const next = window.prompt('Set reminder date (YYYY-MM-DD)', current);
                        if (next == null) { closeMenus(); return; }
                        try { await window.api?.updateDocumentStatus?.(doc.document_id, { reminder_date: next }); refresh(); } catch (err) { window.alert(err?.message || 'Unable to set reminder date'); }
                        closeMenus();
                      }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">Set reminder…</button>
                      <button type="button" onClick={() => { handleOpen(doc); closeMenus(); }} disabled={!doc.file_path} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60">Open</button>
                      <button type="button" onClick={() => { handleReveal(doc); closeMenus(); }} disabled={!doc.file_path} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60">Reveal in Finder</button>
                      <button type="button" onClick={() => { handleMarkPaidToggle(doc); closeMenus(); }} className="w-full text-left rounded px-2 py-1 text-sm text-slate-700 hover:bg-slate-100">
                        {String(status).toLowerCase() === 'paid' ? 'Mark unpaid' : 'Mark paid'}
                      </button>
                      <button type="button" onClick={() => { onDeleteDocument?.(doc); closeMenus(); }} disabled={!doc?.document_id} className="w-full text-left rounded px-2 py-1 text-sm text-red-600 hover:bg-red-50 disabled:opacity-60">Delete</button>
                    </div>
                  ) : null}
                </div>
              </div>
            );
          })}
          {(!loading && computed.length === 0) ? (
            <div className="px-3 py-4 text-sm text-slate-500">No invoices yet.</div>
          ) : null}
        </div>
      </div>
    </div>
    {rebuildOpen ? (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4" role="dialog" aria-modal="true">
        <div className="w-full max-w-2xl rounded-lg bg-white shadow-xl">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <div>
              <h3 className="text-base font-semibold text-slate-800">Rebuild from filename</h3>
              <p className="text-xs text-slate-500">Review proposed changes before applying.</p>
            </div>
            <button className="text-slate-400 hover:text-slate-600" onClick={() => setRebuildOpen(false)} aria-label="Close">✕</button>
          </div>
          <div className="p-4 space-y-3">
            {rebuildLoading ? (
              <div className="text-sm text-slate-500">Analysing filename…</div>
            ) : rebuildPreview && rebuildPreview.ok ? (
              <div className="space-y-3">
                <div className="rounded border border-slate-200">
                  <div className="grid grid-cols-3 gap-2 bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-600">
                    <div>Field</div>
                    <div>Current</div>
                    <div>Proposed</div>
                  </div>
                  {(() => {
                    const rows = [];
                    const current = rebuildDoc || {};
                    const proposed = (rebuildPreview && rebuildPreview.proposed) || {};
                    const add = (label, curVal, nextVal, fmt) => {
                      const format = (v) => (fmt ? fmt(v) : (v == null || v === '' ? '—' : String(v)));
                      const cur = format(curVal);
                      const next = format(nextVal);
                      const changed = cur !== next;
                      rows.push(
                        <div key={label} className="grid grid-cols-3 gap-2 px-3 py-2 text-sm border-t border-slate-100">
                          <div className="text-slate-600">{label}</div>
                          <div className="text-slate-700">{cur}</div>
                          <div className={changed ? 'text-indigo-700 font-medium' : 'text-slate-700'}>{next}</div>
                        </div>
                      );
                    };
                    add('Invoice #', current.number, proposed.number, v => (v == null ? '—' : `INV-${v}`));
                    add('Client', current.client_name || current.display_client_name, proposed.client_name);
                    add('Event name', current.event_name || current.display_event_name, proposed.event_name);
                    add('Event date', current.event_date || current.display_event_date, proposed.event_date);
                    add('Variant', current.invoice_variant, proposed.invoice_variant);
                    add('Amount', current.total_amount ?? current.balance_due, proposed.total_amount, v => v == null ? '—' : toCurrency(v));
                    add('Balance due', current.balance_due, proposed.balance_due, v => v == null ? '—' : toCurrency(v));
                    add('Due date', current.due_date, proposed.due_date);
                    return rows;
                  })()}
                </div>
                {rebuildPreview.matched_jobsheet_id ? (
                  <div className="text-xs text-slate-500">Matched jobsheet ID: {rebuildPreview.matched_jobsheet_id}</div>
                ) : (
                  <div className="text-xs text-slate-500">No strict jobsheet match (using filename tokens only).</div>
                )}
              </div>
            ) : (
              <div className="text-sm text-red-600">{(rebuildPreview && rebuildPreview.message) || 'Unable to preview changes.'}</div>
            )}
          </div>
          <div className="flex items-center justify-end gap-2 border-t border-slate-200 px-4 py-3">
            <button type="button" className="rounded border border-slate-300 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50" onClick={() => setRebuildOpen(false)}>Cancel</button>
            <button
              type="button"
              disabled={!rebuildPreview || !rebuildPreview.ok || rebuildLoading}
              className="rounded bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white hover:bg-indigo-500 disabled:opacity-60"
              onClick={async () => {
                try {
                  setRebuildLoading(true);
                  await window.api?.rebuildInvoiceFromFilename?.({ businessId, documentId: rebuildDoc?.document_id, preview: false });
                  setRebuildOpen(false);
                  setRebuildPreview(null);
                  setRebuildDoc(null);
                  await refresh();
                } catch (err) {
                  window.alert(err?.message || 'Unable to apply rebuild');
                } finally {
                  setRebuildLoading(false);
                }
              }}
            >
              Apply changes
            </button>
          </div>
        </div>
      </div>
    ) : null}
    {relinkOpen ? (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4" role="dialog" aria-modal="true">
        <div className="w-full max-w-4xl rounded-lg bg-white shadow-xl">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <div>
              <h3 className="text-base font-semibold text-slate-800">Relink to jobsheet</h3>
              <p className="text-xs text-slate-500">Choose a jobsheet to link this invoice to, then apply the proposed updates.</p>
            </div>
            <button className="text-slate-400 hover:text-slate-600" onClick={() => setRelinkOpen(false)} aria-label="Close">✕</button>
          </div>
          <div className="p-4 grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <div className="flex items-center gap-2">
                <input type="search" value={relinkSearch} onChange={e=>setRelinkSearch(e.target.value)} placeholder="Search jobsheets" className="w-full rounded border border-slate-300 px-2 py-1 text-sm" />
              </div>
              <div className="rounded border border-slate-200 max-h-80 overflow-auto divide-y divide-slate-100">
                {relinkLoading ? (
                  <div className="px-3 py-2 text-sm text-slate-500">Loading…</div>
                ) : (
                  (relinkJobsheets || [])
                    .filter(js => {
                      const q = relinkSearch.trim().toLowerCase();
                      if (!q) return true;
                      const hay = [js.client_name, js.event_type, js.event_date, js.venue_name].join(' ').toLowerCase();
                      return hay.includes(q);
                    })
                    .map(js => {
                      const id = Number(js.jobsheet_id);
                      const selected = relinkSelectedId != null && Number(relinkSelectedId) === id;
                      return (
                        <button key={id} type="button" onClick={async () => {
                          setRelinkSelectedId(id);
                          setRelinkPreview(null);
                          try {
                            const res = await window.api?.relinkInvoiceToJobsheet?.({ businessId, documentId: relinkDoc?.document_id, jobsheetId: id, preview: true });
                            setRelinkPreview(res || null);
                          } catch (err) {
                            setRelinkPreview({ ok: false, message: err?.message || 'Unable to preview relink' });
                          }
                        }} className={`w-full text-left px-3 py-2 text-sm ${selected ? 'bg-indigo-50' : 'bg-white'} hover:bg-slate-50`}>
                          <div className="font-medium text-slate-800">{js.client_name || 'Untitled'} — {js.event_type || 'Event'}</div>
                          <div className="text-xs text-slate-500">{formatDate(js.event_date)} · {js.venue_name || js.venue_town || ''}</div>
                        </button>
                      );
                    })
                )}
              </div>
            </div>
            <div className="space-y-2">
              <div className="text-sm font-semibold text-slate-700">Proposed updates</div>
              <div className="rounded border border-slate-200 min-h-[6rem]">
                {!relinkPreview ? (
                  <div className="px-3 py-2 text-sm text-slate-500">Select a jobsheet to preview changes.</div>
                ) : relinkPreview.ok ? (
                  <div>
                    <div className="grid grid-cols-3 gap-2 bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-600">
                      <div>Field</div>
                      <div>Current</div>
                      <div>Proposed</div>
                    </div>
                    {(() => {
                      const rows = [];
                      const current = relinkDoc || {};
                      const proposed = (relinkPreview && relinkPreview.proposed) || {};
                      const add = (label, curVal, nextVal, fmt) => {
                        const format = (v) => (fmt ? fmt(v) : (v == null || v === '' ? '—' : String(v)));
                        const cur = format(curVal);
                        const next = format(nextVal);
                        const changed = cur !== next;
                        rows.push(
                          <div key={label} className="grid grid-cols-3 gap-2 px-3 py-2 text-sm border-t border-slate-100">
                            <div className="text-slate-600">{label}</div>
                            <div className="text-slate-700">{cur}</div>
                            <div className={changed ? 'text-indigo-700 font-medium' : 'text-slate-700'}>{next}</div>
                          </div>
                        );
                      };
                      add('Jobsheet ID', current.jobsheet_id, relinkSelectedId, v => v == null ? '—' : String(v));
                      add('Client', current.client_name || current.display_client_name, proposed.client_name);
                      add('Event name', current.event_name || current.display_event_name, proposed.event_name);
                      add('Event date', current.event_date || current.display_event_date, proposed.event_date);
                      add('Variant', current.invoice_variant, proposed.invoice_variant);
                      add('Amount', current.total_amount ?? current.balance_due, proposed.total_amount, v => v == null ? '—' : toCurrency(v));
                      add('Due date', current.due_date, proposed.due_date);
                      return rows;
                    })()}
                  </div>
                ) : (
                  <div className="px-3 py-2 text-sm text-red-600">{relinkPreview.message || 'Unable to preview changes.'}</div>
                )}
              </div>
            </div>
          </div>
          <div className="flex items-center justify-end gap-2 border-t border-slate-200 px-4 py-3">
            <button type="button" className="rounded border border-slate-300 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50" onClick={() => setRelinkOpen(false)}>Cancel</button>
            <button type="button" disabled={!relinkSelectedId || !relinkPreview || !relinkPreview.ok || relinkLoading} className="rounded bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white hover:bg-indigo-500 disabled:opacity-60" onClick={async () => {
              try {
                setRelinkLoading(true);
                await window.api?.relinkInvoiceToJobsheet?.({ businessId, documentId: relinkDoc?.document_id, jobsheetId: relinkSelectedId, preview: false });
                setRelinkOpen(false);
                setRelinkDoc(null);
                setRelinkSelectedId(null);
                setRelinkPreview(null);
                await refresh();
              } catch (err) {
                window.alert(err?.message || 'Unable to relink invoice');
              } finally { setRelinkLoading(false); }
            }}>Apply relink</button>
          </div>
        </div>
      </div>
    ) : null}
  </>);
}

function OpenIcon({ className = 'h-4 w-4' }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M8 16.5v1.25A1.25 1.25 0 0 0 9.25 19h8A1.75 1.75 0 0 0 19 17.25v-8A1.25 1.25 0 0 0 17.75 8H16.5" />
      <path d="M7 17H6.25A1.25 1.25 0 0 1 5 15.75v-8A1.75 1.75 0 0 1 6.75 6h8A1.25 1.25 0 0 1 16 7.25V8" />
      <path d="M10 14.25 17.25 7" />
      <path d="M13 6h4.75V10.75" />
    </svg>
  );
}

function RevealIcon({ className = 'h-4 w-4' }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M2.25 12s2.75-6.75 9.75-6.75 9.75 6.75 9.75 6.75-2.75 6.75-9.75 6.75S2.25 12 2.25 12Z" />
      <path d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z" />
    </svg>
  );
}

function DeleteIcon({ className = 'h-4 w-4' }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M6 7h12" />
      <path d="M9.5 7V5.75A1.75 1.75 0 0 1 11.25 4h1.5A1.75 1.75 0 0 1 14.5 5.75V7" />
      <path d="M17 7v10.25A1.75 1.75 0 0 1 15.25 19h-6.5A1.75 1.75 0 0 1 7 17.25V7" />
      <path d="M10 11v5" />
      <path d="M14 11v5" />
    </svg>
  );
}

function FolderGlyph({ className = 'h-4 w-4' }) {
  return (
    <svg className={className} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 6.25A2.25 2.25 0 0 1 6.25 4h4.086c.414 0 .812.165 1.105.459L13.5 6.5H19A2 2 0 0 1 21 8.5V9H4V6.25Z" fill="currentColor" opacity="0.5" />
      <path d="M3 9.75A1.75 1.75 0 0 1 4.75 8h15.5A1.75 1.75 0 0 1 22 9.75v7.5A2.75 2.75 0 0 1 19.25 20H6A3 3 0 0 1 3 17V9.75Z" fill="currentColor" />
    </svg>
  );
}

function FileGlyph({ className = 'h-4 w-4' }) {
  return (
    <svg className={className} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 3a2 2 0 0 1 2-2h4.172a2 2 0 0 1 1.414.586l4.828 4.828A2 2 0 0 1 20 7.828V20a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V3Z" fill="currentColor" opacity="0.35" />
      <path d="M13 3.5v2.75A1.75 1.75 0 0 0 14.75 8h2.75" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M8.75 12h6.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M8.75 15.5h6.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function DocumentNodeIcon({ isDirectory }) {
  const iconClasses = isDirectory ? 'text-yellow-500' : 'text-slate-500';

  return (
    <span className={`flex h-6 w-6 flex-shrink-0 items-center justify-center ${iconClasses}`} aria-hidden="true">
      {isDirectory ? <FolderGlyph className="h-5 w-5" /> : <FileGlyph className="h-5 w-5" />}
    </span>
  );
}

function DocumentTreeView({
  root,
  trash,
  rootPath,
  loading,
  error,
  onRefresh,
  onOpen,
  onReveal,
  onDeleteFolder,
  onDeleteFile,
  onEmptyTrash,
  emptyingTrash,
  isConfigured,
  collapsed,
  onCollapsedChange,
  persist = false,
  persistKey = ''
}) {
  const safeRootPath = rootPath || '';
  const rootLabel = safeRootPath ? (safeRootPath.split(/[\\/]+/).filter(Boolean).pop() || 'Documents') : 'Documents';
  const [collapsedNodes, setCollapsedNodes] = useState(() => new Set());
  const [selectedNodeId, setSelectedNodeId] = useState(null);
  const isControlledCollapse = typeof collapsed === 'boolean';
  const [internalCollapsed, setInternalCollapsed] = useState(() => (isControlledCollapse ? collapsed : false));

  useEffect(() => {
    if (isControlledCollapse) return;
    if (typeof collapsed === 'boolean') {
      setInternalCollapsed(collapsed);
    }
  }, [collapsed, isControlledCollapse]);

  const panelCollapsed = isControlledCollapse ? collapsed : internalCollapsed;

  const handleSetPanelCollapsed = useCallback((next) => {
    const value = Boolean(next);
    if (!isControlledCollapse) {
      setInternalCollapsed(value);
    }
    onCollapsedChange?.(value);
  }, [isControlledCollapse, onCollapsedChange]);

  useEffect(() => {
    const prefix = persist && persistKey ? `${persistKey}:tree:` : '';
    if (persist && prefix && typeof window !== 'undefined') {
      try {
        const savedCollapsed = window.localStorage.getItem(`${prefix}collapsed`);
        if (savedCollapsed) {
          const arr = JSON.parse(savedCollapsed);
          if (Array.isArray(arr)) setCollapsedNodes(new Set(arr));
        } else {
          setCollapsedNodes(new Set());
        }
        const savedSelected = window.localStorage.getItem(`${prefix}selected`);
        setSelectedNodeId(savedSelected || null);
      } catch (_err) {
        setCollapsedNodes(new Set());
        setSelectedNodeId(null);
      }
    } else {
      setCollapsedNodes(new Set());
      setSelectedNodeId(null);
    }
  }, [safeRootPath, root, persist, persistKey]);

  const toggleFolder = useCallback((nodeId) => {
    setCollapsedNodes(prev => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      if (persist && persistKey && typeof window !== 'undefined') {
        try { window.localStorage.setItem(`${persistKey}:tree:collapsed`, JSON.stringify(Array.from(next))); } catch (_err) {}
      }
      return next;
    });
  }, [persist, persistKey]);

  const handleNodeDoubleClick = useCallback((node, nodeId, isDirectory) => {
    if (!node) return;
    if (isDirectory) {
      toggleFolder(nodeId);
      return;
    }
    if (onOpen) {
      onOpen(node);
    } else if (onReveal) {
      onReveal(node);
    }
  }, [onOpen, onReveal, toggleFolder]);

  const renderRows = (node, depth = 0, isRoot = false) => {
    if (!node) return [];
    const isDirectory = node.isDirectory !== false;
    const absolutePath = node.absolutePath || '';
    const nodeName = node.name || (isRoot ? rootLabel : '(unnamed)');
    const nodeId = isRoot ? '__root__' : (node.path || absolutePath || nodeName || `${nodeName}-${depth}`);
    const hasChildren = Array.isArray(node.children) && node.children.length > 0;
    const isExpanded = isRoot ? true : !collapsedNodes.has(nodeId);
    const isSelected = selectedNodeId === nodeId;

    const rows = [];
    const rowKey = absolutePath || `${nodeName}-${depth}`;
    const itemCount = isDirectory ? Number(node.itemCount || (node.children ? node.children.length : 0)) : 1;
    const sizeValue = isDirectory ? node.totalSize ?? node.size : node.size;
    const modifiedLabel = formatTimestampDisplay(node.modified);

    const baseRowClass = isDirectory ? 'bg-indigo-50/50' : 'bg-white';
    const rowClasses = isSelected
      ? 'bg-indigo-200/80 font-semibold text-indigo-900'
      : `${baseRowClass} hover:bg-indigo-100/70`;

    rows.push(
      <tr
        key={rowKey}
        className={`group cursor-default border-b border-indigo-100 last:border-b-0 transition ${rowClasses}`}
        onClick={() => {
          setSelectedNodeId(nodeId);
          if (persist && persistKey && typeof window !== 'undefined') {
            try { window.localStorage.setItem(`${persistKey}:tree:selected`, nodeId); } catch (_err) {}
          }
        }}
        onDoubleClick={(event) => {
          event.stopPropagation();
          handleNodeDoubleClick(node, nodeId, isDirectory);
        }}
        aria-selected={isSelected}
      >
        <td className="px-3 py-2 text-sm text-slate-700">
          <div className="flex min-w-0 items-center gap-2" style={{ paddingLeft: `${depth * 1.25}rem` }}>
            {isDirectory && hasChildren ? (
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  toggleFolder(nodeId);
                }}
                className="flex h-6 w-6 items-center justify-center rounded text-indigo-400 transition hover:bg-indigo-100 hover:text-indigo-600"
                aria-label={isExpanded ? 'Collapse folder' : 'Expand folder'}
                aria-expanded={isExpanded}
              >
                <span>{isExpanded ? '▾' : '▸'}</span>
              </button>
            ) : (
              <span className="h-6 w-6" />
            )}
            <DocumentNodeIcon isDirectory={isDirectory} />
            <div className="min-w-0 truncate" title={absolutePath || nodeName}>{nodeName}</div>
          </div>
        </td>
        <td className="px-3 py-2 text-xs">
          <span className={`inline-flex items-center rounded-full px-2 py-0.5 font-semibold ${isDirectory ? 'bg-indigo-100 text-indigo-700' : 'bg-sky-100 text-sky-700'}`}>
            {isDirectory ? 'Folder' : 'File'}
          </span>
        </td>
        <td className={`px-3 py-2 text-xs ${isDirectory ? 'text-indigo-600 font-medium' : 'text-slate-500'}`}>
          {isDirectory ? `${itemCount} item${itemCount === 1 ? '' : 's'}` : '—'}
        </td>
        <td className="px-3 py-2 text-xs text-slate-600 font-medium">{formatFileSize(sizeValue)}</td>
        <td className="px-3 py-2 text-xs text-slate-600">{modifiedLabel}</td>
        <td className="px-3 py-2">
          <div className="flex justify-end gap-1 text-indigo-600">
            <TreeActionButton title="Open" onClick={() => onOpen?.(node)} disabled={!onOpen || !absolutePath}>
              <OpenIcon className="h-3.5 w-3.5" />
            </TreeActionButton>
            <TreeActionButton title="Reveal in Finder" onClick={() => onReveal?.(node)} disabled={!onReveal || !absolutePath}>
              <RevealIcon className="h-3.5 w-3.5" />
            </TreeActionButton>
            {!isRoot ? (
              <TreeActionButton
                title="Move to trash"
                onClick={() => {
                  if (isDirectory) {
                    onDeleteFolder?.(node);
                  } else {
                    onDeleteFile?.(node);
                  }
                }}
                disabled={(isDirectory && !onDeleteFolder) || (!isDirectory && !onDeleteFile)}
              >
                <DeleteIcon className="h-3.5 w-3.5" />
              </TreeActionButton>
            ) : null}
          </div>
        </td>
      </tr>
    );

    if (isDirectory && hasChildren && isExpanded) {
      for (const child of node.children) {
        rows.push(...renderRows(child, depth + 1, false));
      }
    }

    return rows;
  };

  let content = null;
  if (panelCollapsed) {
    content = null;
  } else if (!isConfigured) {
    content = (
      <div className="rounded border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-700">
        Choose a documents folder in Settings to enable file management.
      </div>
    );
  } else if (error) {
    content = (
      <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-700" role="alert">
        {error}
      </div>
    );
  } else if (loading) {
    content = <div className="text-xs text-slate-500">Loading tree…</div>;
  } else if (!root || !Array.isArray(root.children) || root.children.length === 0) {
    content = <div className="text-xs text-slate-500">No documents found in this folder yet.</div>;
  } else {
    const tableRows = renderRows(root, 0, true);
    content = (
      <div className="max-h-[420px] overflow-y-auto rounded border border-slate-200">
        <table className="min-w-full border-collapse text-left text-sm">
          <thead className="bg-indigo-600 text-indigo-50">
            <tr>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-xs font-semibold uppercase tracking-wide">Name</th>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-xs font-semibold uppercase tracking-wide">Type</th>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-xs font-semibold uppercase tracking-wide">Contents</th>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-xs font-semibold uppercase tracking-wide">Size</th>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-xs font-semibold uppercase tracking-wide">Modified</th>
              <th scope="col" className="sticky top-0 z-10 bg-indigo-600 px-3 py-2 text-right text-xs font-semibold uppercase tracking-wide">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-indigo-50">{tableRows}</tbody>
        </table>
      </div>
    );
  }

  const trashSummary = trash && typeof trash === 'object' ? trash : null;
  const trashItems = Number.isFinite(trashSummary?.itemCount) ? trashSummary.itemCount : 0;
  const trashSizeLabel = Number.isFinite(trashSummary?.size) ? formatFileSize(trashSummary.size) : null;

  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="text-sm font-semibold text-slate-700">Documents tree</h3>
          <p className="text-xs text-slate-500 break-all">{safeRootPath || 'Not configured'}</p>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => handleSetPanelCollapsed(!panelCollapsed)}
            className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
            aria-expanded={!panelCollapsed}
          >
            {panelCollapsed ? 'Expand' : 'Collapse'}
          </button>
          <button
            type="button"
            onClick={() => onRefresh?.()}
            disabled={loading || !isConfigured}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {loading ? 'Loading…' : 'Refresh'}
          </button>
        </div>
      </div>
      {content}
      {!panelCollapsed && isConfigured ? (
        <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600">
          <div className="flex items-center justify-between gap-3">
            <div>
              <div className="font-semibold text-slate-600">Trash</div>
              <div className="text-xs text-slate-500">
                {trashItems ? `${trashItems} item${trashItems === 1 ? '' : 's'}${trashSizeLabel ? ` · ${trashSizeLabel}` : ''}` : 'Trash is empty'}
              </div>
            </div>
            <button
              type="button"
              onClick={() => onEmptyTrash?.()}
              disabled={!trashItems || emptyingTrash}
              className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {emptyingTrash ? 'Emptying…' : 'Empty trash'}
            </button>
          </div>
        </div>
      ) : null}
      {panelCollapsed ? (
        <div className="text-xs text-slate-500">Tree hidden. Use Expand to view folders.</div>
      ) : null}
    </div>
  );
}

function addDays(dateStr, offset) {
  if (!dateStr) return '';
  const base = new Date(dateStr);
  if (Number.isNaN(base.valueOf())) return '';
  base.setDate(base.getDate() + offset);
  return base.toISOString().slice(0, 10);
}

function toCurrency(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return '£0.00';
  return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(amount);
}

function parseSelectedSingers(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return normalizeSingerEntries(raw);
  try {
    const parsed = JSON.parse(raw);
    return normalizeSingerEntries(parsed);
  } catch (_) {
    return [];
  }
}

function parseProductionItems(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return normalizeProductionItems(raw);
  try {
    const parsed = JSON.parse(raw);
    return normalizeProductionItems(parsed);
  } catch (_) {
    return [];
  }
}

function preparePayload(formState, businessId) {
  const payload = { ...formState, business_id: businessId };

  if (Array.isArray(payload.pricing_selected_singers)) {
    payload.pricing_selected_singers = JSON.stringify(
      normalizeSingerEntries(payload.pricing_selected_singers)
    );
  }

  if (Array.isArray(payload.pricing_production_items)) {
    payload.pricing_production_items = JSON.stringify(
      normalizeProductionItems(payload.pricing_production_items)
    );
  }

  if (payload.pricing_discount_type) {
    payload.pricing_discount_type = String(payload.pricing_discount_type);
  }

  AHMEN_NUMERIC_FIELDS.forEach(field => {
    if (!(field in payload)) return;
    const numeric = Number(payload[field]);
    payload[field] = Number.isFinite(numeric) ? numeric : null;
  });

  AHMEN_BOOLEAN_FIELDS.forEach(field => {
    if (!(field in payload)) return;
    payload[field] = payload[field] ? 1 : 0;
  });

  if (!payload.venue_id) payload.venue_id = null;

  return payload;
}

function applyDerivedFields(nextState) {
  const next = { ...nextState };

  const singerFee = Number(next.ahmen_fee);
  const productionSource = next.pricing_production_total ?? next.production_fees;
  const productionFee = Number(productionSource);
  const totalForDeposit = [singerFee, productionFee]
    .map(amount => (Number.isFinite(amount) && amount > 0 ? amount : 0))
    .reduce((sum, value) => sum + value, 0);

  if (totalForDeposit > 0) {
    const deposit = Math.round(totalForDeposit * 0.3 * 100) / 100;
    const balance = Math.max(totalForDeposit - deposit, 0);
    next.deposit_amount = deposit.toFixed(2);
    next.balance_amount = balance.toFixed(2);
  } else {
    next.deposit_amount = '';
    next.balance_amount = '';
  }

  if (next.pricing_production_total !== undefined) {
    const productionString = next.pricing_production_total ? String(next.pricing_production_total) : '';
    next.production_fees = productionString;
  }

  if (next.event_date) {
    next.balance_due_date = addDays(next.event_date, -10);
    next.balance_reminder_date = addDays(next.event_date, -20);
  } else {
    next.balance_due_date = '';
    next.balance_reminder_date = '';
  }

  if (next.venue_same_as_client) {
    // Force override: ignore/clear any saved venue selection
    next.venue_id = null;
    next.venue_name = 'Private residence';
    next.venue_address1 = next.client_address1 || '';
    next.venue_address2 = next.client_address2 || '';
    next.venue_address3 = next.client_address3 || '';
    next.venue_town = next.client_town || '';
    next.venue_postcode = next.client_postcode || '';
  }

  return next;
}

// BusinessChooser removed – single-business app

function JobsheetList({
  business,
  jobsheets,
  onOpen,
  onNew,
  onDelete,
  onStatusChange,
  onArchiveToggle,
  includeArchived = false,
  onToggleIncludeArchived,
  loading,
  deletingId,
  statusUpdatingId,
  sortConfig,
  onSort,
  activeJobsheetId = null
}) {
  const [searchValue, setSearchValue] = useState('');
  const [statusFilters, setStatusFilters] = useState(() => new Set());
  const normalizedSearch = searchValue.trim().toLowerCase();
  const [exportingPersonnel, setExportingPersonnel] = useState(false);
  const [exportPanelOpen, setExportPanelOpen] = useState(false);
  const [exportFormat, setExportFormat] = useState(() => {
    try { const raw = window.localStorage.getItem(storageKey); if (raw) { const p = JSON.parse(raw); return p.format || 'pdf'; } } catch (_) {}
    return 'pdf';
  });
  const defaultCols = ['date','time','client','event','venue','personnel'];
  const storageKey = useMemo(() => business ? `ui:${business.id}:personnelExport` : 'ui:personnelExport', [business]);
  const [exportFromDate, setExportFromDate] = useState(() => {
    try { const raw = window.localStorage.getItem(storageKey); if (raw) { const p = JSON.parse(raw); return p.fromDate || ''; } } catch (_) {}
    return '';
  });
  const [exportToDate, setExportToDate] = useState(() => {
    try { const raw = window.localStorage.getItem(storageKey); if (raw) { const p = JSON.parse(raw); return p.toDate || ''; } } catch (_) {}
    return '';
  });
  const [exportCols, setExportCols] = useState(() => {
    try { const raw = window.localStorage.getItem(storageKey); if (raw) { const p = JSON.parse(raw); if (Array.isArray(p.columns) && p.columns.length) return p.columns; } } catch (_) {}
    return defaultCols;
  });
  useEffect(() => {
    try { window.localStorage.setItem(storageKey, JSON.stringify({ fromDate: exportFromDate, toDate: exportToDate, columns: exportCols, format: exportFormat })); } catch (_) {}
  }, [storageKey, exportFromDate, exportToDate, exportCols, exportFormat]);

  const handleExportPersonnel = useCallback(async () => {
    try {
      if (!business || !business.id) return;
      setExportingPersonnel(true);
      const payload = { businessId: business.id };
      if (exportFromDate) payload.fromDate = exportFromDate;
      if (exportToDate) payload.toDate = exportToDate;
      if (Array.isArray(exportCols) && exportCols.length) payload.columns = exportCols;
      if (exportFormat === 'text') {
        const res = await window.api?.createPersonnelLogText?.(payload);
        if (!res || res.ok !== true || !res.text) throw new Error(res?.message || 'Unable to create personnel text');
        try { await window.api?.copyTextToClipboard?.(res.text); } catch (_) {}
        window.alert('Personnel list copied to clipboard');
      } else {
        const res = await window.api?.createPersonnelLogPdf?.(payload);
        if (!res || res.ok !== true) throw new Error(res?.message || 'Unable to create personnel PDF');
        try { await window.api?.showItemInFolder?.(res.file_path); } catch (_) {}
        window.alert('Personnel log saved to:\n' + (res.file_path || ''));
      }
    } catch (err) {
      window.alert(err?.message || 'Unable to export personnel log');
    } finally {
      setExportingPersonnel(false);
      setExportPanelOpen(false);
    }
  }, [business, exportFromDate, exportToDate, exportCols, exportFormat]);

  const allColumnOptions = [
    { key: 'date', label: 'Date' },
    { key: 'time', label: 'Time' },
    { key: 'status', label: 'Status' },
    { key: 'client', label: 'Client' },
    { key: 'event', label: 'Event' },
    { key: 'venue', label: 'Venue' },
    { key: 'personnel', label: 'Personnel' },
    { key: 'singer_count', label: 'Singer count' },
    { key: 'total', label: 'Total (est.)' },
    { key: 'notes', label: 'Notes' }
  ];
  const toggleExportCol = useCallback((key) => {
    setExportCols(prev => {
      const set = new Set(prev);
      if (set.has(key)) set.delete(key); else set.add(key);
      const next = Array.from(set);
      // Maintain a sensible ordering based on allColumnOptions
      const order = allColumnOptions.map(o => o.key);
      next.sort((a, b) => order.indexOf(a) - order.indexOf(b));
      return next;
    });
  }, []);

  // Column controls: show/hide + reorder (persist per business)
  const JOBSHEET_COLUMNS_STORAGE_KEY = `ui:${business?.id}:jobsheetColumns`;
  const defaultOrder = useMemo(() => JOBSHEET_COLUMNS.map(c => c.key), []);
  const [columnsMenuOpen, setColumnsMenuOpen] = useState(false);
  const [columnsMenuAbove, setColumnsMenuAbove] = useState(false);
  const columnsMenuRef = useRef(null);
  const columnsMenuContentRef = useRef(null);
  const [columnOrder, setColumnOrder] = useState(() => {
    if (typeof window === 'undefined') return defaultOrder;
    try {
      const raw = window.localStorage.getItem(JOBSHEET_COLUMNS_STORAGE_KEY);
      if (!raw) return defaultOrder;
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.order)) return parsed.order.filter(Boolean);
    } catch (_) {}
    return defaultOrder;
  });
  const [columnVisibility, setColumnVisibility] = useState(() => {
    if (typeof window === 'undefined') return {};
    try {
      const raw = window.localStorage.getItem(JOBSHEET_COLUMNS_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (parsed && parsed.visibility && typeof parsed.visibility === 'object') return parsed.visibility;
    } catch (_) {}
    return {};
  });
  useEffect(() => {
    if (typeof window === 'undefined' || !business) return;
    try {
      window.localStorage.setItem(JOBSHEET_COLUMNS_STORAGE_KEY, JSON.stringify({ order: columnOrder, visibility: columnVisibility }));
    } catch (_) {}
  }, [business, columnOrder, columnVisibility]);
  const baseColumnMap = useMemo(() => new Map(JOBSHEET_COLUMNS.map(c => [c.key, c])), []);
  const effectiveColumns = useMemo(() => {
    const normalizedOrder = [...columnOrder].filter(k => baseColumnMap.has(k));
    for (const key of baseColumnMap.keys()) if (!normalizedOrder.includes(key)) normalizedOrder.push(key);
    const list = normalizedOrder
      .map(k => baseColumnMap.get(k))
      .filter(Boolean)
      .filter(col => col.key === 'actions' ? true : (columnVisibility[col.key] !== false));
    // Always keep actions last if present
    const others = list.filter(c => c.key !== 'actions');
    const actions = list.find(c => c.key === 'actions');
    return actions ? [...others, actions] : others;
  }, [columnOrder, columnVisibility, baseColumnMap]);

  const moveColumn = useCallback((key, dir) => {
    setColumnOrder(prev => {
      const arr = prev.slice();
      const idx = arr.indexOf(key);
      if (idx < 0) return arr;
      const swapWith = dir === 'up' ? idx - 1 : idx + 1;
      if (swapWith < 0 || swapWith >= arr.length) return arr;
      const tmp = arr[swapWith];
      arr[swapWith] = arr[idx];
      arr[idx] = tmp;
      return arr;
    });
  }, []);
  const toggleColumn = useCallback((key) => {
    if (key === 'actions') return; // cannot hide actions
    setColumnVisibility(prev => {
      const currentVisible = prev?.[key] !== false; // default visible
      const nextVisible = !currentVisible;
      return { ...prev, [key]: nextVisible ? true : false };
    });
  }, []);
  useEffect(() => {
    if (!columnsMenuOpen) return undefined;
    const onDoc = (e) => {
      if (columnsMenuRef.current && !columnsMenuRef.current.contains(e.target)) setColumnsMenuOpen(false);
    };
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, [columnsMenuOpen]);

  const filteredJobsheets = useMemo(() => {
    if (!jobsheets || jobsheets.length === 0) {
      return [];
    }

    const activeStatuses = Array.from(statusFilters);
    const hasStatusFilter = activeStatuses.length > 0;

    return jobsheets.filter(sheet => {
      if (!sheet) return false;
      if (hasStatusFilter && !activeStatuses.includes(sheet.status)) {
        return false;
      }

      if (!normalizedSearch) {
        return true;
      }

      const formattedEventDate = sheet.event_date ? formatDateDisplay(sheet.event_date) : '';
      const haystack = [
        sheet.jobsheet_id != null ? `#${sheet.jobsheet_id}` : '',
        sheet.client_name,
        sheet.client_email,
        sheet.client_phone,
        sheet.event_type,
        sheet.event_date,
        formattedEventDate,
        sheet.venue_name,
        sheet.venue_town,
        sheet.venue_postcode,
        sheet.venue_address1,
        sheet.venue_address2,
        sheet.venue_address3,
        sheet.notes
      ];

      return haystack.some(value => {
        if (value == null || value === '') return false;
        return String(value).toLowerCase().includes(normalizedSearch);
      });
    });
  }, [jobsheets, normalizedSearch, statusFilters]);

  const sortedJobsheets = useMemo(() => {
    const list = [...filteredJobsheets];
    const { key, direction } = sortConfig || {};
    if (!key) return list;
    const multiplier = direction === 'asc' ? 1 : -1;

    const getComparableValue = (sheet, field) => {
      switch (field) {
        case 'event_date':
          return sheet.event_date ? new Date(sheet.event_date).valueOf() : 0;
        case 'ahmen_fee':
          {
            const total = Number(sheet.pricing_total);
            if (Number.isFinite(total) && total > 0) return total;
            const singerFee = Number(sheet.ahmen_fee) || 0;
            const productionFee = Number(sheet.production_fees) || 0;
            return singerFee + productionFee;
          }
        case 'status':
          return STATUS_ORDER[sheet.status] ?? STATUS_OPTIONS.length;
        case 'client_name':
        case 'event_type':
          return (sheet[field] || '').toString().toLowerCase();
        case 'venue_name':
          return (sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '').toString().toLowerCase();
        default:
          return sheet[field];
      }
    };

    list.sort((a, b) => {
      const valueA = getComparableValue(a, key);
      const valueB = getComparableValue(b, key);

      if (valueA === valueB) return 0;

      if (typeof valueA === 'number' && typeof valueB === 'number') {
        return (valueA - valueB) * multiplier;
      }

      return String(valueA ?? '').localeCompare(String(valueB ?? ''), 'en', { sensitivity: 'base' }) * multiplier;
    });

    return list;
  }, [filteredJobsheets, sortConfig]);

  const toggleStatusFilter = useCallback((status) => {
    if (!status) return;
    setStatusFilters(prev => {
      const next = new Set(prev);
      if (next.has(status)) {
        next.delete(status);
      } else {
        next.add(status);
      }
      return next;
    });
  }, []);

  const clearFilters = useCallback(() => {
    setSearchValue('');
    setStatusFilters(() => new Set());
  }, []);

  const hasActiveFilters = Boolean(normalizedSearch) || statusFilters.size > 0;
  const totalCount = jobsheets?.length || 0;
  const filteredCount = filteredJobsheets.length;

  const summaryLabel = hasActiveFilters
    ? `${filteredCount} of ${totalCount} jobsheets`
    : `${filteredCount} jobsheet${filteredCount === 1 ? '' : 's'}`;

  const renderSortIndicator = (columnKey) => {
    if (!sortConfig || sortConfig.key !== columnKey) return <span className="text-slate-400 ml-1">⇅</span>;
    return (
      <span className="ml-1 text-xs text-indigo-600">
        {sortConfig.direction === 'asc' ? '▲' : '▼'}
      </span>
    );
  };

  const renderHeaderRow = () => (
    <div
      className="grid items-center gap-3 rounded-lg bg-slate-50 px-3 py-2"
      style={{ gridTemplateColumns: JOBSHEET_GRID_TEMPLATE }}
    >
      {JOBSHEET_COLUMNS.map(column => {
        const baseClass = `${getAlignmentClasses(column.align)} flex w-full items-center`;
        const labelClass = `${baseClass} text-xs font-semibold uppercase tracking-wide text-slate-600`;
        if (!column.sortable) {
          return (
            <div key={column.key} className={labelClass}>
              {column.label}
            </div>
          );
        }
        return (
          <button
            key={column.key}
            type="button"
            onClick={() => onSort(column.key)}
            className={`${labelClass} gap-1 bg-transparent p-0 border-0 hover:text-indigo-600 focus:outline-none focus:ring-0`}
          >
            {column.label}
            {renderSortIndicator(column.key)}
          </button>
        );
      })}
    </div>
  );

  const renderDataRow = (sheet) => {
    const statusKey = normalizeStatus(sheet.status) || 'enquiry';
    const statusStyles = STATUS_STYLES[statusKey] || 'bg-slate-200 text-slate-700 border border-slate-300';
    const statusDisabled = statusUpdatingId === sheet.jobsheet_id;
    const baseRowClass = STATUS_ROW_CLASSES[statusKey] || 'bg-white';
    const activeRowClass = ACTIVE_STATUS_ROW_CLASSES[statusKey] || baseRowClass;
    const numericRowId = sheet.jobsheet_id != null ? Number(sheet.jobsheet_id) : null;
    const isActive = numericRowId != null && activeJobsheetId != null && Number(activeJobsheetId) === numericRowId;
    const isArchived = Boolean(sheet.archived_at);

    const rowBackground = isActive ? activeRowClass : baseRowClass;
    const baseCellClass = 'px-4 py-3 text-sm';
    const verticalBorder = 'border-y border-transparent';
    const firstCellExtras = isActive
      ? "relative before:absolute before:inset-y-2 before:left-1 before:w-1 before:rounded-full before:bg-indigo-600 before:content-[''] before:block rounded-l-xl shadow-[0_0_0_2px_rgba(79,70,229,0.25)]"
      : 'rounded-l-xl';
    const lastCellExtras = isActive
      ? 'rounded-r-xl shadow-[0_0_0_2px_rgba(79,70,229,0.25)]'
      : 'rounded-r-xl';

    return (
      <tr
        key={sheet.jobsheet_id || sheet.client_name}
        onClick={() => onOpen(sheet.jobsheet_id)}
        className={`cursor-pointer ${isArchived ? 'opacity-70' : ''}`}
      >
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} ${firstCellExtras}`}>
          <div className="flex items-center gap-3 min-w-0">
            {isActive ? <span className="h-8 w-1 rounded-full bg-indigo-600" /> : null}
            <span className="font-medium text-slate-800 whitespace-normal break-words">{sheet.client_name || 'Untitled booking'}</span>
          </div>
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder}`}>
          {sheet.event_type || '—'}
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} whitespace-nowrap`}>
          {formatDateDisplay(sheet.event_date)}
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder}`}>
          <div className="min-w-0 whitespace-normal break-words">
            {sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '—'}
          </div>
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder}`}>
          <div className="flex justify-center">
            <select
              value={statusKey}
              disabled={statusDisabled}
              className={`rounded-full px-3 py-1 text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500 ${statusStyles} ${statusDisabled ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'}`}
              onClick={event => event.stopPropagation()}
              onMouseDown={event => event.stopPropagation()}
              onChange={event => {
                event.stopPropagation();
                const nextStatus = event.target.value;
                if (!nextStatus || nextStatus === statusKey) return;
                onStatusChange?.(sheet.jobsheet_id, nextStatus);
              }}
            >
              {STATUS_OPTIONS.map(option => (
                <option key={option.value} value={option.value}>{option.label}</option>
              ))}
            </select>
          </div>
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} text-right text-slate-600`}>
          {toCurrency((Number(sheet.pricing_total) || (Number(sheet.ahmen_fee) || 0) + (Number(sheet.production_fees) || 0)))}
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} ${lastCellExtras}`}>
          <div className="flex justify-end gap-2">
            <button
              type="button"
              onClick={(event) => { event.stopPropagation(); onArchiveToggle?.(sheet.jobsheet_id, !isArchived); }}
              className={`rounded border px-2 py-1 text-xs font-medium ${isArchived ? 'border-slate-200 text-slate-600 hover:bg-slate-50' : 'border-amber-200 text-amber-700 hover:bg-amber-50'}`}
            >
              {isArchived ? 'Unarchive' : 'Archive'}
            </button>
            <button
              type="button"
              disabled={deletingId === sheet.jobsheet_id}
              onClick={(event) => {
                event.stopPropagation();
                onDelete(sheet.jobsheet_id);
              }}
              className="rounded border border-red-200 px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-50 disabled:opacity-60"
            >
              Delete
            </button>
          </div>
        </td>
      </tr>
    );
  };

  return (
    <div className="flex flex-col h-full">
      <div className="mb-4 space-y-3">
        <div className="flex min-w-0 flex-col gap-3 sm:flex-row sm:flex-wrap sm:items-center sm:justify-between">
          <div>
            <h2 className="text-lg font-semibold text-slate-700">Jobsheets</h2>
            <p className="text-sm text-slate-500">{summaryLabel}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <ImportJobsheetButton business={business} onCreated={(id) => onOpen?.(id)} />
            <div className="relative inline-block">
              <button
                type="button"
                onClick={() => setExportPanelOpen(v => !v)}
                className="inline-flex items-center gap-2 rounded border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:opacity-60"
              >
                {exportingPersonnel ? 'Exporting…' : 'Export Personnel PDF'}
              </button>
              {exportPanelOpen && (
                <div className="absolute right-0 z-50 mt-2 w-80 rounded border border-slate-200 bg-white p-3 text-sm shadow-lg">
                  <div className="mb-2 font-medium text-slate-700">Customize</div>
                  <div className="mb-2 grid grid-cols-2 gap-x-3 gap-y-2">
                    {allColumnOptions.map(opt => (
                      <label key={opt.key} className="inline-flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={exportCols.includes(opt.key)}
                          onChange={() => toggleExportCol(opt.key)}
                        />
                        <span>{opt.label}</span>
                      </label>
                    ))}
                  </div>
                  <div className="mb-2 grid grid-cols-2 gap-3">
                    <label className="block">
                      <div className="text-xs text-slate-500">From date</div>
                      <input type="date" value={exportFromDate} onChange={e => setExportFromDate(e.target.value)} className="w-full rounded border border-slate-300 px-2 py-1" />
                    </label>
                    <label className="block">
                      <div className="text-xs text-slate-500">To date</div>
                      <input type="date" value={exportToDate} onChange={e => setExportToDate(e.target.value)} className="w-full rounded border border-slate-300 px-2 py-1" />
                    </label>
                  </div>
                  <div className="mb-2">
                    <div className="text-xs text-slate-500 mb-1">Format</div>
                    <div className="inline-flex items-center gap-4">
                      <label className="inline-flex items-center gap-1">
                        <input type="radio" name="exportFormat" value="pdf" checked={exportFormat === 'pdf'} onChange={() => setExportFormat('pdf')} />
                        <span>PDF</span>
                      </label>
                      <label className="inline-flex items-center gap-1">
                        <input type="radio" name="exportFormat" value="text" checked={exportFormat === 'text'} onChange={() => setExportFormat('text')} />
                        <span>Text (copy to WhatsApp)</span>
                      </label>
                    </div>
                  </div>
                  <div className="flex items-center justify-end gap-2">
                    <button type="button" onClick={() => setExportPanelOpen(false)} className="rounded border border-slate-200 px-2 py-1 text-xs text-slate-600 hover:bg-slate-50">Cancel</button>
                    <button type="button" onClick={handleExportPersonnel} disabled={exportingPersonnel || exportCols.length === 0} className="rounded bg-indigo-600 px-3 py-1 text-xs font-medium text-white hover:bg-indigo-500 disabled:opacity-60">Export</button>
                  </div>
                </div>
              )}
            </div>
            <button
              onClick={onNew}
              className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-medium px-3 py-2 rounded"
            >
              + New Jobsheet
            </button>
          </div>
        </div>
        <div className="flex flex-col gap-3 md:flex-row md:flex-wrap md:items-center md:justify-between">
              <div className="relative w-full md:max-w-xs">
                <input
                  type="search"
                  value={searchValue}
                  onChange={event => setSearchValue(event.target.value)}
                  placeholder="Search jobsheets"
                  className="w-full rounded border border-slate-300 bg-white px-3 py-2 text-sm text-slate-700 shadow-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                />
              </div>
          <div className="flex flex-wrap items-center gap-2">
            {STATUS_OPTIONS.map(option => {
              const active = statusFilters.has(option.value);
              return (
                <button
                  key={option.value}
                  type="button"
                  onClick={() => toggleStatusFilter(option.value)}
                  className={`inline-flex items-center rounded-full border px-3 py-1 text-xs font-medium transition ${active ? 'border-indigo-200 bg-indigo-50 text-indigo-700 shadow-sm' : 'border-slate-200 bg-white text-slate-500 hover:border-slate-300 hover:bg-slate-50'}`}
                  aria-pressed={active}
                >
                  {option.label}
                </button>
              );
            })}
            {hasActiveFilters ? (
              <button
                type="button"
                onClick={clearFilters}
                className="inline-flex items-center rounded-full border border-slate-200 px-3 py-1 text-xs font-medium text-slate-500 hover:bg-slate-50"
              >
                Clear
              </button>
            ) : null}
            <label className="inline-flex items-center gap-2 text-xs text-slate-600 ml-2">
              <input type="checkbox" checked={!!includeArchived} onChange={() => onToggleIncludeArchived?.()} />
              Show archived
            </label>
            <div className="relative" ref={columnsMenuRef}>
              <button
                type="button"
                onClick={() => setColumnsMenuOpen(v => !v)}
                className="inline-flex items-center rounded-full border px-3 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
              >
                Columns
              </button>
              {columnsMenuOpen ? (
                <div
                  ref={columnsMenuContentRef}
                  className={`absolute right-0 z-20 w-56 rounded border border-slate-200 bg-white p-2 shadow-lg ${columnsMenuAbove ? 'bottom-full mb-2' : 'top-full mt-2'}`}
                >
                  {JOBSHEET_COLUMNS.map(col => (
                    <div key={col.key} className="flex items-center justify-between gap-2 px-1 py-1 text-sm">
                      <label className="inline-flex items-center gap-2 text-slate-700">
                        <input type="checkbox" disabled={col.key === 'actions'} checked={col.key === 'actions' ? true : columnVisibility[col.key] !== false} onChange={() => toggleColumn(col.key)} />
                        <span>{col.label || (col.key === 'actions' ? 'Actions' : col.key)}</span>
                      </label>
                      <div className="ml-2 flex items-center gap-1">
                        <button type="button" className="rounded border border-slate-300 px-1 text-xs text-slate-600 hover:bg-slate-50" onClick={() => moveColumn(col.key, 'up')}>↑</button>
                        <button type="button" className="rounded border border-slate-300 px-1 text-xs text-slate-600 hover:bg-slate-50" onClick={() => moveColumn(col.key, 'down')}>↓</button>
                      </div>
                    </div>
                  ))}
                </div>
              ) : null}
            </div>
          </div>
        </div>
      </div>
      <div className="flex-1 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {loading ? (
          <div className="p-6 text-center text-slate-500">Loading…</div>
        ) : !sortedJobsheets.length ? (
          <div className="p-6 text-center text-slate-500">{hasActiveFilters ? 'No jobsheets match your filters yet. Adjust the search or status filter to see more results.' : 'No jobsheets yet. Create your first one!'}</div>
        ) : (
          <div className="overflow-y-auto overflow-x-auto max-h-[55vh]">
            <table className="min-w-full text-sm border-separate border-spacing-y-2">
              <thead>
                <tr className="bg-slate-50">
                  {effectiveColumns.map(column => {
                    const alignClass = column.align === 'right'
                      ? 'text-right'
                      : column.align === 'center'
                        ? 'text-center'
                        : 'text-left';
                    if (!column.sortable) {
                      return (
                        <th
                          key={column.key}
                          scope="col"
                          className={`px-4 py-2 text-xs font-semibold uppercase tracking-wide text-slate-600 ${alignClass}`}
                        >
                          {column.label}
                        </th>
                      );
                    }
                    return (
                      <th key={column.key} scope="col" className={`px-4 py-2 text-xs font-semibold uppercase tracking-wide text-slate-600 ${alignClass}`}>
                        <button
                          type="button"
                          onClick={() => onSort(column.key)}
                          className="inline-flex items-center gap-1 bg-transparent p-0 text-slate-600 hover:text-indigo-600 focus:outline-none focus:ring-0"
                        >
                          {column.label}
                          {renderSortIndicator(column.key)}
                        </button>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {sortedJobsheets.map(sheet => {
                  const statusKey = normalizeStatus(sheet.status) || 'enquiry';
                  const statusStyles = STATUS_STYLES[statusKey] || 'bg-slate-200 text-slate-700 border border-slate-300';
                  const statusDisabled = statusUpdatingId === sheet.jobsheet_id;
                  const baseRowClass = STATUS_ROW_CLASSES[statusKey] || 'bg-white';
                  const activeRowClass = ACTIVE_STATUS_ROW_CLASSES[statusKey] || baseRowClass;
                  const numericRowId = sheet.jobsheet_id != null ? Number(sheet.jobsheet_id) : null;
                  const isActive = numericRowId != null && activeJobsheetId != null && Number(activeJobsheetId) === numericRowId;
                  const rowBackground = isActive ? activeRowClass : baseRowClass;
                  const baseCellClass = 'px-4 py-3 text-sm';
                  const verticalBorder = 'border-y border-transparent';
                  const firstCellExtras = isActive ? "relative before:absolute before:inset-y-2 before:left-1 before:w-1 before:rounded-full before:bg-indigo-600 before:content-[''] before:block rounded-l-xl shadow-[0_0_0_2px_rgba(79,70,229,0.25)]" : 'rounded-l-xl';
                  const lastCellExtras = isActive ? 'rounded-r-xl shadow-[0_0_0_2px_rgba(79,70,229,0.25)]' : 'rounded-r-xl';
                  const currency = toCurrency((Number(sheet.pricing_total) || (Number(sheet.ahmen_fee) || 0) + (Number(sheet.production_fees) || 0)));
                  const isArchived = Boolean(sheet.archived_at);
                  return (
                    <tr key={sheet.jobsheet_id || sheet.client_name} onClick={() => onOpen(sheet.jobsheet_id)} className={`cursor-pointer ${isArchived ? 'opacity-70' : ''}`}>
                      {effectiveColumns.map((col, idx) => {
                        const alignClass = col.align === 'right' ? 'text-right' : (col.align === 'center' ? 'text-center' : 'text-left');
                        const isFirst = idx === 0;
                        const isLast = idx === effectiveColumns.length - 1;
                        const extra = isFirst ? firstCellExtras : (isLast ? lastCellExtras : '');
                        const common = `${rowBackground} ${baseCellClass} ${verticalBorder} ${extra}`;
                        switch (col.key) {
                          case 'client_name':
                            return (
                              <td key={col.key} className={`${common} ${alignClass}`}>
                                <div className="flex items-center gap-3 min-w-0">
                                  {isActive ? <span className="h-8 w-1 rounded-full bg-indigo-600" /> : null}
                                  <span className="font-medium text-slate-800 whitespace-normal break-words">{sheet.client_name || 'Untitled booking'}</span>
                                </div>
                              </td>
                            );
                          case 'event_type':
                            return (<td key={col.key} className={`${common} ${alignClass}`}>{sheet.event_type || '—'}</td>);
                          case 'event_date':
                            return (<td key={col.key} className={`${common} ${alignClass} whitespace-nowrap`}>{formatDateDisplay(sheet.event_date)}</td>);
                          case 'venue_name':
                            return (
                              <td key={col.key} className={`${common} ${alignClass}`}>
                                <div className="min-w-0 whitespace-normal break-words">{sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '—'}</div>
                              </td>
                            );
                          case 'status':
                            return (
                              <td key={col.key} className={`${common} ${alignClass}`}>
                                <div className="flex flex-wrap justify-center">
                                  <select
                                    value={statusKey}
                                    disabled={statusDisabled}
                                    className={`rounded-full px-3 py-1 text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500 ${statusStyles} ${statusDisabled ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'}`}
                                    onClick={event => event.stopPropagation()}
                                    onMouseDown={event => event.stopPropagation()}
                                    onChange={event => {
                                      event.stopPropagation();
                                      const nextStatus = event.target.value;
                                      if (!nextStatus || nextStatus === statusKey) return;
                                      onStatusChange?.(sheet.jobsheet_id, nextStatus);
                                    }}
                                  >
                                    {STATUS_OPTIONS.map(option => (
                                      <option key={option.value} value={option.value}>{option.label}</option>
                                    ))}
                                  </select>
                                </div>
                              </td>
                            );
                          case 'ahmen_fee':
                            return (<td key={col.key} className={`${common} ${alignClass} text-slate-600`}>{currency}</td>);
                          case 'actions':
                            return (
                              <td key={col.key} className={`${common} ${alignClass}`}>
                                <div className="flex flex-wrap justify-end gap-2">
                                  <button type="button" onClick={(event) => { event.stopPropagation(); onArchiveToggle?.(sheet.jobsheet_id, !isArchived); }} className={`rounded border px-2 py-1 text-xs font-medium ${isArchived ? 'border-slate-200 text-slate-600 hover:bg-slate-50' : 'border-amber-200 text-amber-700 hover:bg-amber-50'}`}>{isArchived ? 'Unarchive' : 'Archive'}</button>
                                  <button type="button" disabled={deletingId === sheet.jobsheet_id} onClick={(event) => { event.stopPropagation(); onDelete(sheet.jobsheet_id); }} className="rounded border border-red-200 px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-50 disabled:opacity-60">Delete</button>
                                </div>
                              </td>
                            );
                          default:
                            return (<td key={col.key} className={`${common} ${alignClass}`}>—</td>);
                        }
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function ImportJobsheetButton({ business, onCreated }) {
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [errorLocal, setErrorLocal] = useState('');
  const [source, setSource] = useState({ folder: '', workbook_path: '', invoices: [] });
  const [draft, setDraft] = useState({
    // Client
    client_name: '',
    client_email: '',
    client_phone: '',
    client_address1: '',
    client_address2: '',
    client_address3: '',
    client_town: '',
    client_postcode: '',
    // Event
    event_type: '',
    event_date: '',
    event_start: '',
    event_end: '',
    // Venue
    venue_name: '',
    venue_address1: '',
    venue_address2: '',
    venue_address3: '',
    venue_town: '',
    venue_postcode: '',
    // Services
    service_types: '',
    specialist_singers: '',
    caterer_name: ''
  });

  const setField = (key, value) => setDraft(prev => ({ ...prev, [key]: value }));

  // Normalize typed time strings to 24-hour HH:MM
  const to24h = (h, min, ap) => {
    let hour = Number(h);
    let m = Number(min);
    if (Number.isNaN(hour)) hour = 0;
    if (Number.isNaN(m)) m = 0;
    hour = Math.max(0, Math.min(23, hour));
    m = Math.max(0, Math.min(59, m));
    if (ap) {
      const ampm = ap.toUpperCase();
      hour = hour % 12;
      if (ampm === 'PM') hour += 12;
    }
    return `${String(hour).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  };
  const normalizeTime24 = (input) => {
    const raw = (input || '').toString().trim();
    if (!raw) return '';
    let s = raw.replace(/\./g, ':').replace(/-/g, ':').replace(/\s+/g, ' ').trim();
    // 12h with optional minutes and optional space before am/pm
    let m = s.match(/^(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])$/);
    if (m) return to24h(m[1], m[2] ?? '0', m[3]);
    // 24h HH:MM
    m = s.match(/^(\d{1,2}):(\d{1,2})$/);
    if (m) return to24h(m[1], m[2]);
    // Compact 3-4 digits e.g. 730 or 1530
    m = s.match(/^(\d{3,4})$/);
    if (m) {
      const num = m[1];
      const mm = num.slice(-2);
      const hh = num.slice(0, num.length - 2);
      return to24h(hh, mm);
    }
    // Bare hour
    m = s.match(/^(\d{1,2})$/);
    if (m) return to24h(m[1], '0');
    return raw;
  };

  useEffect(() => {
    // Lock background scroll when modal is open
    try {
      if (open) {
        const prev = document.body.style.overflow;
        document.body.dataset.prevOverflow = prev || '';
        document.body.style.overflow = 'hidden';
      } else if (document.body.dataset.prevOverflow !== undefined) {
        document.body.style.overflow = document.body.dataset.prevOverflow;
        delete document.body.dataset.prevOverflow;
      }
    } catch (_) {}
    return () => {
      try {
        if (document.body.dataset.prevOverflow !== undefined) {
          document.body.style.overflow = document.body.dataset.prevOverflow;
          delete document.body.dataset.prevOverflow;
        }
      } catch (_) {}
    };
  }, [open]);

  const handleChoose = async () => {
    try {
      setErrorLocal('');
      const folder = await window.api?.chooseDirectory?.({ title: 'Select source for import', defaultPath: business?.save_path || undefined });
      if (!folder) return;
      setLoading(true);
      const res = await window.api?.extractJobsheetFromFolder?.({ folderPath: folder });
      setLoading(false);
      if (!res || res.ok === false) { setErrorLocal(res?.message || 'Unable to extract'); return; }
      const sug = res.suggested || {};
      const init = {
        client_name: sug.client_name || '',
        client_email: sug.client_email || '',
        client_phone: sug.client_phone || '',
        client_address1: sug.client_address1 || '',
        client_address2: sug.client_address2 || '',
        client_address3: sug.client_address3 || '',
        client_town: sug.client_town || '',
        client_postcode: sug.client_postcode || '',
        event_type: sug.event_type || '',
        event_date: sug.event_date || '',
        event_start: sug.event_start || '',
        event_end: sug.event_end || '',
        venue_name: sug.venue_name || '',
        venue_address1: sug.venue_address1 || '',
        venue_address2: sug.venue_address2 || '',
        venue_address3: sug.venue_address3 || '',
        venue_town: sug.venue_town || '',
        venue_postcode: sug.venue_postcode || '',
        service_types: sug.service_types || '',
        specialist_singers: sug.specialist_singers || '',
        caterer_name: sug.caterer_name || ''
      };
      setDraft(init);
      setSource({ folder: res.folder || folder, workbook_path: res.workbook_path || '', invoices: Array.isArray(res.invoices) ? res.invoices : [] });
      setOpen(true);
    } catch (err) {
      setLoading(false);
      console.error('Import failed', err);
      setErrorLocal(err?.message || 'Import failed');
    }
  };

  const handleApply = async () => {
    try {
      setErrorLocal('');
      const client = (draft.client_name || '').trim();
      if (!client) { setErrorLocal('Client name is required'); return; }
      const payload = {
        business_id: business?.id,
        status: 'contracting',
        client_name: client,
        event_date: draft.event_date || null,
        client_email: draft.client_email || null,
        client_phone: draft.client_phone || null,
        client_address1: draft.client_address1 || null,
        client_address2: draft.client_address2 || null,
        client_address3: draft.client_address3 || null,
        client_town: draft.client_town || null,
        client_postcode: draft.client_postcode || null,
        event_type: draft.event_type || null,
        event_start: draft.event_start || null,
        event_end: draft.event_end || null,
        venue_name: draft.venue_name || null,
        venue_address1: draft.venue_address1 || null,
        venue_address2: draft.venue_address2 || null,
        venue_address3: draft.venue_address3 || null,
        venue_town: draft.venue_town || null,
        venue_postcode: draft.venue_postcode || null,
        service_types: draft.service_types || null,
        specialist_singers: draft.specialist_singers || null,
        caterer_name: draft.caterer_name || null
      };
      const id = await window.api?.addAhmenJobsheet?.(payload);
      if (id) {
        window.api?.notifyJobsheetChange?.({ type: 'jobsheet-created', businessId: business?.id, jobsheetId: id });
        setOpen(false);
        onCreated?.(id);
      } else {
        setErrorLocal('Failed to create jobsheet');
      }
    } catch (err) {
      console.error('Create from import failed', err);
      setErrorLocal(err?.message || 'Unable to create jobsheet');
    }
  };

  return (
    <>
      <button
        type="button"
        className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
        onClick={handleChoose}
        disabled={loading}
      >
        {loading ? 'Scanning…' : 'Import'}
      </button>

      {open ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4">
          <div
            className="w-full max-w-2xl rounded-lg bg-white p-6 shadow-xl overflow-y-auto"
            style={{ maxHeight: '60vh' }}
          >
            <div className="flex items-start justify-between border-b border-slate-200 pb-4">
              <div>
                <h3 className="text-lg font-semibold text-slate-800">Import</h3>
                <p className="text-sm text-slate-500">Review and edit values before creating the jobsheet.</p>
              </div>
              <button type="button" onClick={() => setOpen(false)} className="text-slate-400 hover:text-slate-600" aria-label="Close import modal">✕</button>
            </div>
            <div className="mt-4 space-y-4">
              {errorLocal ? (
                <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-600">{errorLocal}</div>
              ) : null}
              <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600 flex items-center justify-between gap-2">
                <div className="truncate">
                  <div><span className="font-medium">Source:</span> <span className="truncate" title={source.folder}>{source.folder || '—'}</span></div>
                  <div><span className="font-medium">Workbook:</span> <span className="truncate" title={source.workbook_path}>{source.workbook_path || '—'}</span></div>
                  <div><span className="font-medium">Invoices:</span> {source.invoices && source.invoices.length ? `${source.invoices.length} found` : 'none'}</div>
                </div>
                <div className="flex items-center gap-2">
                  {source.workbook_path ? (
                    <button type="button" className="rounded border border-slate-300 px-2 py-1 text-xs" onClick={() => window.api?.openPath?.(source.workbook_path)}>Open workbook</button>
                  ) : null}
                  {source.folder ? (
                    <button type="button" className="rounded border border-slate-300 px-2 py-1 text-xs" onClick={() => window.api?.openPath?.(source.folder)}>Open folder</button>
                  ) : null}
                </div>
              </div>

              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 mb-2">Client</div>
                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-sm font-medium text-slate-600">
                  Client name
                  <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_name} onChange={e => setField('client_name', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Email
                    <input type="email" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_email} onChange={e => setField('client_email', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Phone
                    <input type="tel" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_phone} onChange={e => setField('client_phone', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 1
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_address1} onChange={e => setField('client_address1', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 2
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_address2} onChange={e => setField('client_address2', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 3
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_address3} onChange={e => setField('client_address3', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Town / City
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_town} onChange={e => setField('client_town', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Postcode
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.client_postcode} onChange={e => setField('client_postcode', e.target.value)} />
                  </label>
                </div>
              </div>

              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 mb-2">Event</div>
                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-sm font-medium text-slate-600">
                  Event date
                  <input type="date" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={formatDateInput(draft.event_date)} onChange={e => setField('event_date', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Event type
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.event_type} onChange={e => setField('event_type', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Start time
                    <input
                      type="text"
                      placeholder="e.g. 19:00"
                      className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      value={draft.event_start}
                      onChange={e => setField('event_start', e.target.value)}
                      onBlur={e => setField('event_start', normalizeTime24(e.target.value))}
                    />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    End time
                    <input
                      type="text"
                      placeholder="e.g. 22:30"
                      className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      value={draft.event_end}
                      onChange={e => setField('event_end', e.target.value)}
                      onBlur={e => setField('event_end', normalizeTime24(e.target.value))}
                    />
                  </label>
                </div>
              </div>

              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 mb-2">Venue</div>
                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-sm font-medium text-slate-600">
                    Venue name
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_name} onChange={e => setField('venue_name', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 1 (venue)
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_address1} onChange={e => setField('venue_address1', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 2 (venue)
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_address2} onChange={e => setField('venue_address2', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Address line 3 (venue)
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_address3} onChange={e => setField('venue_address3', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Town / City
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_town} onChange={e => setField('venue_town', e.target.value)} />
                  </label>
                  <label className="text-sm font-medium text-slate-600">
                    Postcode
                    <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.venue_postcode} onChange={e => setField('venue_postcode', e.target.value)} />
                  </label>
                </div>
              </div>

              

              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 mb-2">Services & Notes</div>
                <label className="text-sm font-medium text-slate-600">
                  Service types / ensemble
                  <textarea rows={2} className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.service_types} onChange={e => setField('service_types', e.target.value)} />
                </label>
                <label className="text-sm font-medium text-slate-600">
                  Specialist singers
                  <textarea rows={2} className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.specialist_singers} onChange={e => setField('specialist_singers', e.target.value)} />
                </label>
                <label className="text-sm font-medium text-slate-600">
                  Caterer name
                  <input type="text" className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" value={draft.caterer_name} onChange={e => setField('caterer_name', e.target.value)} />
                </label>
              </div>

              <p className="text-[11px] text-slate-500">You can adjust or add missing fields in the editor after creation.</p>
            </div>
            <div className="mt-4 pt-4 border-t border-slate-200 flex items-center justify-end gap-3">
              <button type="button" onClick={() => setOpen(false)} className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">Cancel</button>
              <button
                type="button"
                className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500"
                onClick={handleApply}
              >
                Create jobsheet
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}


function InlineJobsheetEditorPanel({
  business,
  visible,
  jobsheetId,
  sessionKey,
  onClose,
  onOpenInWindow
}) {
  const headerTitle = jobsheetId ? 'Edit jobsheet' : 'New jobsheet';
  const hint = jobsheetId
    ? 'Changes save automatically.'
    : 'Changes save automatically. Fill in the details below.';

  if (!visible) {
    return (
      <div className="mx-auto max-w-7xl">
        <div className="rounded-lg border border-dashed border-slate-300 bg-white px-6 py-8 text-sm text-slate-500">
          Select a jobsheet from the list (or create a new one) to edit it inline. You can still pop the editor into its own window when needed.
        </div>
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl">
      <div className="rounded-lg border border-slate-200 bg-slate-100 shadow-sm">
        <div className="flex flex-col gap-3 border-b border-slate-200 bg-slate-50 px-5 py-4 sm:flex-row sm:items-center sm:justify-between lg:px-6 lg:py-4">
          <div>
            <h3 className="text-base font-semibold text-slate-700">{headerTitle}</h3>
            <p className="text-xs text-slate-500">{hint}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={onOpenInWindow}
              className="inline-flex items-center gap-1 rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:border-indigo-200 hover:text-indigo-600 focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              Open in window
            </button>
            <button
              type="button"
              onClick={onClose}
              className="inline-flex items-center gap-1 rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:border-red-200 hover:text-red-600 focus:outline-none focus:ring-2 focus:ring-red-500"
            >
              Close editor
            </button>
          </div>
        </div>
        <JobsheetEditorWindow
          key={sessionKey}
          variant="inline"
          businessId={business.id}
          businessName={business.business_name}
          initialJobsheetId={jobsheetId == null ? 'new' : jobsheetId}
          targetJobsheetId={jobsheetId}
          onRequestClose={onClose}
        />
      </div>
    </div>
  );
}


function normalizeSingerEntries(entries) {
  const list = Array.isArray(entries) ? entries : [];
  const seen = new Set();
  const normalized = [];
  list.forEach(entry => {
    if (entry == null) return;
    let id;
    let fee = '';
    let name = '';
    let custom = false;
    if (typeof entry === 'string') {
      id = entry;
    } else if (typeof entry === 'object') {
      id = entry.id ?? entry.singerId ?? entry.value;
      if (entry.fee !== undefined && entry.fee !== null) {
        fee = entry.fee === '' ? '' : String(entry.fee);
      }
      name = entry.name ?? entry.label ?? entry.title ?? '';
      custom = Boolean(entry.custom);
    }
    if (!id) return;
    const key = String(id);
    if (seen.has(key) && !custom) return;
    seen.add(key);
    const normalizedEntry = { id: key, fee };
    if (name) normalizedEntry.name = String(name);
    if (custom) normalizedEntry.custom = true;
    if (typeof entry === 'object' && entry && entry.locked === true) normalizedEntry.locked = true;
    normalized.push(normalizedEntry);
  });
  return normalized;
}

function equalSingerEntries(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (a[i].id !== b[i].id) return false;
    const feeA = a[i].fee ?? '';
    const feeB = b[i].fee ?? '';
    if (String(feeA) !== String(feeB)) return false;
    const nameA = a[i].name ?? '';
    const nameB = b[i].name ?? '';
    if (nameA !== nameB) return false;
    const customA = Boolean(a[i].custom);
    const customB = Boolean(b[i].custom);
    if (customA !== customB) return false;
    const lockedA = Boolean(a[i].locked);
    const lockedB = Boolean(b[i].locked);
    if (lockedA !== lockedB) return false;
    // availability flag removed
  }
  return true;
}



function PricingPanel({ pricingConfig, formState, onChange, pricingTotals, hasExisting = false, onUpdateSingerPool, onFocusPricingPanel }) {
  if (!pricingConfig) {
    return (
      <div className="rounded border border-slate-200 bg-white p-4 text-sm text-slate-500">
        Loading pricing configuration…
      </div>
    );
  }

  const serviceTypes = pricingConfig.serviceTypes ?? [];
  // Availability tracking removed; UI simplified
  const existingPool = Array.isArray(pricingConfig.singerPool) ? pricingConfig.singerPool : [];
  const singerPool = useMemo(
    () => existingPool.map(singer => ({
      ...singer,
      id: singer?.id != null ? String(singer.id) : ''
    })).filter(singer => singer.id),
    [existingPool]
  );

  const sortedSingers = useMemo(() => {
    const currentServiceId = formState.pricing_service_id != null ? String(formState.pricing_service_id) : '';
    return [...singerPool].sort((a, b) => {
      const aDefault = currentServiceId
        ? Boolean(a.serviceFees?.[currentServiceId]?.defaultIncluded)
        : Boolean(a.defaultIncluded);
      const bDefault = currentServiceId
        ? Boolean(b.serviceFees?.[currentServiceId]?.defaultIncluded)
        : Boolean(b.defaultIncluded);
      if (aDefault !== bDefault) return aDefault ? -1 : 1;
      return (a.name || '').localeCompare(b.name || '', 'en', { sensitivity: 'base' });
    });
  }, [singerPool, formState.pricing_service_id]);

  const poolMap = useMemo(
    () => new Map(sortedSingers.map(singer => [singer.id, singer])),
    [sortedSingers]
  );

  const canManagePool = typeof onUpdateSingerPool === 'function';

  const selectedEntries = useMemo(
    () => normalizeSingerEntries(formState.pricing_selected_singers),
    [formState.pricing_selected_singers]
  );

  const productionItems = useMemo(
    () => normalizeProductionItems(formState.pricing_production_items),
    [formState.pricing_production_items]
  );

  const productionTotalValue = useMemo(
    () => calculateProductionTotal(productionItems),
    [productionItems]
  );

  const customFeesNumber = Number(formState.pricing_custom_fees) || 0;

  const formatFeeForInput = useCallback((value) => {
    if (value === null || value === undefined || value === '') return '';
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return String(value);
    const fixed = numeric.toFixed(2);
    if (fixed.endsWith('00')) return String(Math.round(numeric));
    return fixed.endsWith('0') ? fixed.slice(0, -1) : fixed;
  }, []);

  const updateSelected = useCallback((entries) => {
    const normalized = normalizeSingerEntries(entries);
    if (!equalSingerEntries(normalized, selectedEntries)) {
      onChange('pricing_selected_singers', normalized);
    }
  }, [onChange, selectedEntries]);

  const selectedServiceId = formState.pricing_service_id != null ? String(formState.pricing_service_id) : '';
  // Preference: auto-select default team when service changes
  const [autoSelectDefaultTeam, setAutoSelectDefaultTeam] = useState(() => {
    try { return window.localStorage.getItem('pricing:autoDefaultTeam') !== '0'; } catch (_) { return true; }
  });
  useEffect(() => {
    try { window.localStorage.setItem('pricing:autoDefaultTeam', autoSelectDefaultTeam ? '1' : '0'); } catch (_) {}
  }, [autoSelectDefaultTeam]);
  const selectedService = serviceTypes.find(type => String(type.id) === selectedServiceId) || null;
  const lastServiceIdRef = useRef('');
  const serviceEffectDidMountRef = useRef(false);

  useEffect(() => {
    const currentServiceId = selectedService ? String(selectedService.id) : '';
    // Avoid re-applying default singers on the initial mount if a service is already selected
    // (e.g., navigating away and back). Still allow defaults when the user actively changes service after mount.
    if (!serviceEffectDidMountRef.current) {
      serviceEffectDidMountRef.current = true;
      if (currentServiceId) {
        lastServiceIdRef.current = currentServiceId;
        return;
      }
    }
    if (!currentServiceId) {
      if (selectedEntries.length) updateSelected([]);
      lastServiceIdRef.current = '';
      return;
    }

    if (currentServiceId !== lastServiceIdRef.current) {
      lastServiceIdRef.current = currentServiceId;
      if (!autoSelectDefaultTeam) {
        // Respect user preference: do not auto-apply defaults on service change
        return;
      }
      const previousMap = new Map(selectedEntries.map(e => [e.id, e]));
      const defaults = sortedSingers
        .filter(singer => {
          const serviceDetails = singer.serviceFees?.[currentServiceId];
          if (serviceDetails) return Boolean(serviceDetails.defaultIncluded);
          return Boolean(singer.defaultIncluded);
        })
        .map(singer => {
          const serviceDetails = singer.serviceFees?.[currentServiceId];
          const fallbackFee = singer.fee != null ? String(singer.fee) : '';
          const fee = serviceDetails?.fee != null ? String(serviceDetails.fee) : fallbackFee;
          const prev = previousMap.get(singer.id);
          if (prev && prev.locked) {
            return { ...prev };
          }
          return {
            id: singer.id,
            name: singer.name,
            fee
          };
        });
      // Carry over any previously selected entries that aren't in defaults (custom or non-default picks), preserving locked ones
      const defaultIds = new Set(defaults.map(d => d.id));
      const carry = selectedEntries.filter(e => !defaultIds.has(e.id));
      updateSelected([...defaults, ...carry]);
      return;
    }

    const normalized = selectedEntries
      .map(entry => {
        const poolSinger = poolMap.get(entry.id);
        if (!poolSinger) {
          return entry.custom ? entry : null;
        }
        if (entry.locked) {
          return entry;
        }
        const serviceDetails = poolSinger.serviceFees?.[currentServiceId];
        const fallbackFee = poolSinger.fee != null ? String(poolSinger.fee) : '';
        const fee = entry.fee !== undefined && entry.fee !== ''
          ? String(entry.fee)
          : serviceDetails?.fee != null ? String(serviceDetails.fee) : fallbackFee;
        return {
          ...entry,
          name: poolSinger.name || entry.name,
          fee
        };
      })
      .filter(Boolean);

    if (!equalSingerEntries(normalized, selectedEntries)) {
      updateSelected(normalized);
    }
  }, [selectedService, sortedSingers, selectedEntries, poolMap, updateSelected, autoSelectDefaultTeam]);

  const internalTotals = useMemo(() => {
    let base = 0;
    let singerCount = 0;
    selectedEntries.forEach(entry => {
      const singer = poolMap.get(entry.id);
      if (singer) {
        const feeValue = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
          ? Number(entry.fee)
          : Number(singer.fee);
        base += Number.isFinite(feeValue) ? feeValue : 0;
        singerCount += 1;
      } else if (entry.custom) {
        const feeValue = Number(entry.fee);
        base += Number.isFinite(feeValue) ? feeValue : 0;
        singerCount += 1;
      }
    });

    const singerSubtotal = base + customFeesNumber;
    const singerDiscountValue = calculateDiscountValue({
      type: formState.pricing_discount_type || 'amount',
      value: formState.pricing_discount,
      subtotal: singerSubtotal
    });
    const singerNet = Math.max(singerSubtotal - singerDiscountValue, 0);

    const productionDiscountValue = calculateDiscountValue({
      type: formState.pricing_production_discount_type || 'amount',
      value: formState.pricing_production_discount,
      subtotal: productionTotalValue
    });
    const productionNet = Math.max(productionTotalValue - productionDiscountValue, 0);
    const total = Math.max(singerNet + productionNet, 0);
    const hasSelection = singerCount > 0 || customFeesNumber !== 0 || productionTotalValue !== 0 || singerDiscountValue > 0 || productionDiscountValue > 0;
    return {
      base,
      singerCount,
      productionSubtotal: productionTotalValue,
      productionNet,
      productionDiscountValue,
      custom: customFeesNumber,
      singerDiscountValue,
      singerNet,
      subtotal: singerSubtotal + productionTotalValue,
      total,
      hasSelection
    };
  }, [selectedEntries, poolMap, productionTotalValue, customFeesNumber, formState.pricing_discount, formState.pricing_discount_type, formState.pricing_production_discount, formState.pricing_production_discount_type]);

  const totals = pricingTotals || internalTotals;
  const singerDiscountType = formState.pricing_discount_type || 'amount';
  const singerDiscountValueNumber = totals.singerDiscountValue || 0;
  const productionDiscountType = formState.pricing_production_discount_type || 'amount';
  const productionDiscountValueNumber = totals.productionDiscountValue || 0;
  const productionSubtotalValue = totals.productionSubtotal ?? productionTotalValue;
  const productionNetValue = totals.productionNet ?? Math.max(productionSubtotalValue - productionDiscountValueNumber, 0);
  const singerNetValue = totals.singerNet ?? Math.max((totals.base || 0) + customFeesNumber - singerDiscountValueNumber, 0);
  const totalValue = totals.total ?? (singerNetValue + productionNetValue);
  const totalDerivedString = totals.hasSelection ? totalValue.toFixed(2) : '';

  useEffect(() => {
    const nextDiscountString = singerDiscountValueNumber > 0 ? singerDiscountValueNumber.toFixed(2) : '';
    const current = formState.pricing_discount_value ?? '';
    if (nextDiscountString !== current) {
      onChange('pricing_discount_value', nextDiscountString);
    }
  }, [singerDiscountValueNumber, formState.pricing_discount_value, onChange]);

  const selectedIdSet = useMemo(
    () => new Set(selectedEntries.map(entry => entry.id)),
    [selectedEntries]
  );

  const handleToggleSinger = useCallback((singerId) => {
    const poolSinger = poolMap.get(singerId);
    if (!poolSinger) return;
    const serviceId = selectedService ? String(selectedService.id) : '';
    const isSelected = selectedEntries.some(entry => entry.id === singerId);
    if (isSelected) {
      updateSelected(selectedEntries.filter(entry => entry.id !== singerId));
      return;
    }
    const serviceDetails = serviceId ? poolSinger.serviceFees?.[serviceId] : null;
    const fallbackFee = poolSinger.fee != null ? String(poolSinger.fee) : '';
    const fee = serviceDetails?.fee != null ? String(serviceDetails.fee) : fallbackFee;
    updateSelected([
      ...selectedEntries,
      {
        id: singerId,
        name: poolSinger.name,
        fee,
        locked: false
      }
    ]);
  }, [poolMap, selectedEntries, selectedService, updateSelected]);

  const handleClearSelection = useCallback(() => {
    // Preserve locked status in memory by re-adding entries but marked as unselected is equivalent to clearing list.
    // Since selection is represented by presence in the array, we clear the array; locked state will persist when
    // those singers are re-selected (we don't maintain separate memory). If you want to remember locks across clears,
    // we can store a transient map. For now, simply clear selection as requested.
    updateSelected([]);
  }, [updateSelected]);

  const handleSelectDefaultTeam = useCallback(() => {
    const serviceId = selectedService ? String(selectedService.id) : '';
    if (!serviceId) {
      updateSelected([]);
      return;
    }
    const defaults = sortedSingers
      .filter(singer => {
        const serviceDetails = singer.serviceFees?.[serviceId];
        if (serviceDetails) return Boolean(serviceDetails.defaultIncluded);
        return Boolean(singer.defaultIncluded);
      })
      .map(singer => {
        const existing = selectedEntries.find(e => e.id === singer.id);
        if (existing && existing.locked) {
          return { ...existing };
        }
        const serviceDetails = singer.serviceFees?.[serviceId];
        const fallbackFee = singer.fee != null ? String(singer.fee) : '';
        const fee = serviceDetails?.fee != null ? String(serviceDetails.fee) : fallbackFee;
        return {
          id: singer.id,
          name: singer.name,
          fee,
          locked: existing ? Boolean(existing.locked) : false
        };
      });
    // Keep any previously selected entries that aren't part of defaults (e.g., custom or extra picks)
    const defaultIds = new Set(defaults.map(d => d.id));
    const carry = selectedEntries.filter(e => !defaultIds.has(e.id));
    updateSelected([...defaults, ...carry]);
  }, [selectedService, sortedSingers, updateSelected]);

  const currentServiceId = selectedService ? String(selectedService.id) : '';
  const hasDefaultSingers = useMemo(() => {
    if (!currentServiceId) return sortedSingers.some(singer => singer.defaultIncluded);
    return sortedSingers.some(singer => Boolean(singer.serviceFees?.[currentServiceId]?.defaultIncluded));
  }, [sortedSingers, currentServiceId]);

  const [newSingerName, setNewSingerName] = useState('');
  const [newSingerBaseFee, setNewSingerBaseFee] = useState('');
  const [newSingerServiceFees, setNewSingerServiceFees] = useState(() => ({}));
  const [addingSinger, setAddingSinger] = useState(false);
  const [addError, setAddError] = useState('');
  const [showAddSingerModal, setShowAddSingerModal] = useState(false);
  const [showEditSingerModal, setShowEditSingerModal] = useState(false);
  const [editSingerId, setEditSingerId] = useState('');
  const [editSingerName, setEditSingerName] = useState('');
  const [editSingerBaseFee, setEditSingerBaseFee] = useState('');
  const [editSingerServiceFees, setEditSingerServiceFees] = useState({});
  const [editSingerDefaultIncluded, setEditSingerDefaultIncluded] = useState(false);
  const [editingSinger, setEditingSinger] = useState(false);
  const [editError, setEditError] = useState('');

  const resetAddSingerForm = useCallback(() => {
    setNewSingerName('');
    setNewSingerBaseFee('');
    setAddError('');
    setNewSingerServiceFees(() => {
      const initial = {};
      serviceTypes.forEach(service => {
        const serviceId = service.id != null ? String(service.id) : '';
        if (!serviceId) return;
        initial[serviceId] = { fee: '', defaultIncluded: Boolean(service.defaultIncluded) };
      });
      return initial;
    });
  }, [serviceTypes]);

  const handleOpenAddSingerModal = useCallback(() => {
    if (!canManagePool) return;
    resetAddSingerForm();
    setShowAddSingerModal(true);
  }, [canManagePool, resetAddSingerForm]);

  const handleCloseAddSingerModal = useCallback(() => {
    if (addingSinger) return;
    resetAddSingerForm();
    setShowAddSingerModal(false);
  }, [addingSinger, resetAddSingerForm]);

  const handleOpenEditSingerModal = useCallback((singer) => {
    if (!singer || !canManagePool) return;
    setEditSingerId(singer.id);
    setEditSingerName(singer.name || '');
    setEditSingerBaseFee(
      singer.fee !== undefined && singer.fee !== null && singer.fee !== ''
        ? String(singer.fee)
        : ''
    );
    const initialServiceFees = {};
    serviceTypes.forEach(service => {
      const serviceId = service.id != null ? String(service.id) : '';
      const existing = singer.serviceFees?.[serviceId] || null;
      if (!serviceId) return;
      const feeValue = existing?.fee !== undefined && existing?.fee !== null
        ? String(existing.fee)
        : '';
      initialServiceFees[serviceId] = {
        fee: feeValue,
        defaultIncluded: Boolean(existing?.defaultIncluded)
      };
    });
    setEditSingerServiceFees(initialServiceFees);
    setEditSingerDefaultIncluded(Boolean(singer.defaultIncluded));
    setEditError('');
    setShowAddSingerModal(false);
    setShowEditSingerModal(true);
  }, [canManagePool, serviceTypes]);

  const handleCloseEditSingerModal = useCallback(() => {
    if (editingSinger) return;
    setEditSingerId('');
    setEditSingerName('');
    setEditSingerBaseFee('');
    setEditSingerServiceFees({});
    setEditSingerDefaultIncluded(false);
    setEditError('');
    setShowEditSingerModal(false);
  }, [editingSinger]);

  const handleAddSingerToPool = useCallback(async () => {
    if (typeof onUpdateSingerPool !== 'function') return;
    const trimmedName = newSingerName.trim();
    if (!trimmedName) return;

    const baseFeeNumber = Number(newSingerBaseFee);
    const baseFee = newSingerBaseFee === ''
      ? 0
      : Number.isFinite(baseFeeNumber) && baseFeeNumber >= 0 ? baseFeeNumber : 0;

    const serviceFees = {};
    serviceTypes.forEach(service => {
      const serviceId = service.id != null ? String(service.id) : '';
      if (!serviceId) return;
      const config = newSingerServiceFees[serviceId] || { fee: '', defaultIncluded: false };
      const hasFee = config.fee !== '';
      const feeNumber = Number(config.fee);
      if (!hasFee && !config.defaultIncluded) return;
      const normalizedFee = hasFee && Number.isFinite(feeNumber) && feeNumber >= 0
        ? feeNumber
        : baseFee;
      serviceFees[serviceId] = {
        fee: normalizedFee,
        defaultIncluded: Boolean(config.defaultIncluded)
      };
    });

    const nextPool = [
      ...existingPool,
      {
        id: `pool-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        name: trimmedName,
        fee: baseFee,
        defaultIncluded: false,
        serviceFees
      }
    ];

    try {
      setAddingSinger(true);
      setAddError('');
      await onUpdateSingerPool(nextPool);
      resetAddSingerForm();
      setShowAddSingerModal(false);
      onFocusPricingPanel?.();
    } catch (err) {
      console.error('Failed to add singer to pool', err);
      setAddError(err?.message || 'Unable to add singer');
    } finally {
      setAddingSinger(false);
    }
  }, [onUpdateSingerPool, existingPool, newSingerName, newSingerBaseFee, newSingerServiceFees, serviceTypes, resetAddSingerForm]);

  const confirmDisabled = !canManagePool || !newSingerName.trim().length || addingSinger;

  const handleNewSingerServiceFeeChange = useCallback((serviceId, field, value) => {
    setNewSingerServiceFees(prev => ({
      ...prev,
      [serviceId]: {
        fee: field === 'fee' ? value : (prev[serviceId]?.fee ?? ''),
        defaultIncluded: field === 'defaultIncluded'
          ? Boolean(value)
          : Boolean(prev[serviceId]?.defaultIncluded)
      }
    }));
  }, []);

  const handleEditServiceFeeChange = useCallback((serviceId, field, value) => {
    setEditSingerServiceFees(prev => ({
      ...prev,
      [serviceId]: {
        fee: field === 'fee' ? value : (prev[serviceId]?.fee ?? ''),
        defaultIncluded: field === 'defaultIncluded'
          ? Boolean(value)
          : Boolean(prev[serviceId]?.defaultIncluded)
      }
    }));
  }, []);

  const handleSaveEditedSinger = useCallback(async () => {
    if (!canManagePool || !editSingerId.trim()) return;
    const trimmedName = editSingerName.trim();
    if (!trimmedName) {
      setEditError('Name is required');
      return;
    }

    const baseFeeNumber = Number(editSingerBaseFee);
    const normalizedBaseFee = editSingerBaseFee === ''
      ? 0
      : Number.isFinite(baseFeeNumber) && baseFeeNumber >= 0 ? baseFeeNumber : 0;

    const serviceFeeEntries = {};
    Object.entries(editSingerServiceFees).forEach(([serviceId, config]) => {
      if (!serviceId) return;
      const feeNumber = Number(config?.fee);
      const feeValue = config?.fee === ''
        ? undefined
        : Number.isFinite(feeNumber) && feeNumber >= 0 ? feeNumber : undefined;
      if (feeValue === undefined && !config?.defaultIncluded) {
        serviceFeeEntries[serviceId] = {
          fee: undefined,
          defaultIncluded: Boolean(config?.defaultIncluded)
        };
      } else {
        serviceFeeEntries[serviceId] = {
          fee: feeValue !== undefined ? feeValue : normalizedBaseFee,
          defaultIncluded: Boolean(config?.defaultIncluded)
        };
      }
    });

    const nextPool = existingPool.map(singer => (
      singer.id === editSingerId
        ? {
            ...singer,
            name: trimmedName,
            fee: normalizedBaseFee,
            defaultIncluded: Boolean(editSingerDefaultIncluded),
            serviceFees: serviceFeeEntries
          }
        : singer
    ));

    try {
      setEditingSinger(true);
      setEditError('');
      await onUpdateSingerPool(nextPool);
      onFocusPricingPanel?.();
      handleCloseEditSingerModal();
    } catch (err) {
      console.error('Failed to update singer', err);
      setEditError(err?.message || 'Unable to update singer');
    } finally {
      setEditingSinger(false);
    }
  }, [canManagePool, editSingerId, editSingerName, editSingerBaseFee, editSingerServiceFees, editSingerDefaultIncluded, existingPool, onUpdateSingerPool, handleCloseEditSingerModal]);

  const handleDeleteSinger = useCallback(async () => {
    if (!canManagePool || !editSingerId.trim()) return;
    const nextPool = existingPool.filter(singer => singer.id !== editSingerId);
    try {
      setEditingSinger(true);
      setEditError('');
      await onUpdateSingerPool(nextPool);
      onFocusPricingPanel?.();
      handleCloseEditSingerModal();
      const updatedSelection = selectedEntries.filter(entry => entry.id !== editSingerId);
      if (updatedSelection.length !== selectedEntries.length) {
        updateSelected(updatedSelection);
      }
    } catch (err) {
      console.error('Failed to delete singer from pool', err);
      setEditError(err?.message || 'Unable to delete singer');
    } finally {
      setEditingSinger(false);
    }
  }, [canManagePool, editSingerId, existingPool, onUpdateSingerPool, handleCloseEditSingerModal, selectedEntries, updateSelected]);

  return (
    <>
      <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-6">
      <section className="space-y-2">
        <div className="flex items-center justify-between">
          <span className="text-sm font-medium text-slate-600">Service configuration</span>
          {selectedService ? (
            <span className="text-xs text-slate-400">{selectedService.label}</span>
          ) : null}
        </div>
        <div className="mt-2 flex flex-wrap gap-2">
          {serviceTypes.length ? serviceTypes.map(type => {
            const typeId = type.id != null ? String(type.id) : '';
            const isActive = typeId === selectedServiceId;
            return (
              <button
                key={type.id}
                type="button"
                onClick={() => onChange('pricing_service_id', isActive ? '' : type.id)}
                className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-sm font-medium transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${isActive ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:border-indigo-200 hover:text-indigo-600'}`}
              >
                {type.label}
              </button>
            );
          }) : (
            <span className="text-sm text-slate-500">No service templates configured.</span>
          )}
        </div>
      </section>

      <section className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <span className="text-sm font-medium text-slate-600">Select your lineup</span>
          <div className="flex items-center gap-2">
            <label className="inline-flex items-center gap-1.5 text-xs text-slate-600 mr-2">
              <input type="checkbox" checked={autoSelectDefaultTeam} onChange={e => setAutoSelectDefaultTeam(e.target.checked)} />
              Auto-select on service change
            </label>
            <button
              type="button"
              onClick={handleSelectDefaultTeam}
              disabled={!hasDefaultSingers || !selectedServiceId}
              className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              Use default team
            </button>
            <button
              type="button"
              onClick={handleClearSelection}
              disabled={!selectedEntries.length}
              className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              Clear selection
            </button>
            <button
              type="button"
              onClick={handleOpenAddSingerModal}
              disabled={!canManagePool}
              className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              Add singer
            </button>
          </div>
        </div>
        {sortedSingers.length ? (
          <div className="space-y-2">
            {sortedSingers.map(singer => {
              const singerId = singer.id;
              const isSelected = selectedIdSet.has(singerId);
              const serviceDetails = currentServiceId ? singer.serviceFees?.[currentServiceId] : null;
              const displayFee = serviceDetails?.fee != null ? serviceDetails.fee : singer.fee;
              const selectionEntry = selectedEntries.find(entry => entry.id === singerId);
              const isLocked = Boolean(selectionEntry?.locked);
              
              const selectionFee = selectionEntry?.fee;
              const feeInputValue = formatFeeForInput(
                selectionFee !== undefined && selectionFee !== null && selectionFee !== ''
                  ? selectionFee
                  : displayFee
              );

              return (
                <div
                  key={singerId}
                  className={`flex flex-nowrap items-center gap-2 rounded border px-3 py-2 text-sm transition ${
                    isSelected
                      ? isLocked
                        ? 'border-red-300 bg-red-500 text-white shadow-sm'
                        : 'border-indigo-200 bg-indigo-500 text-white shadow-sm'
                      : 'border-slate-200 bg-white text-slate-700 hover:border-indigo-200 hover:bg-indigo-50/70'
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => handleToggleSinger(singerId)}
                    className={`inline-flex h-6 w-6 flex-shrink-0 items-center justify-center rounded-full border text-xs font-semibold focus:outline-none focus:ring-2 ${
                      isSelected ? 'border-white text-white focus:ring-white/60' : 'border-slate-300 text-slate-500 focus:ring-indigo-500'
                    }`}
                    aria-pressed={isSelected}
                  >
                    {isSelected ? '✓' : ''}
                  </button>

                  <div className="flex min-w-[8rem] sm:min-w-[12rem] flex-1 items-center overflow-hidden">
                    <span className={`font-medium leading-tight truncate ${isSelected ? 'text-white' : 'text-slate-700'}`}>
                      {singer.name || 'Unnamed singer'}
                    </span>
                  </div>
                  

                  <label
                    className={`flex w-36 sm:w-44 flex-shrink-0 items-center gap-1 text-xs uppercase tracking-wide ${
                      isSelected ? 'text-white/80' : 'text-slate-500'
                    }`}
                  >
                    <span>Fee</span>
                    <div className="relative flex items-center">
                      <span className={`pointer-events-none absolute left-2 ${isSelected ? 'text-white/70' : 'text-slate-400'}`}>£</span>
                      <input
                        type="number"
                        step="0.01"
                        className={`w-24 sm:w-28 rounded border px-5 py-1 text-sm focus:outline-none focus:ring-2 ${
                          isSelected
                            ? 'border-white/70 bg-white text-indigo-700 placeholder-indigo-300 focus:ring-white/60'
                            : 'border-slate-300 bg-white text-slate-700 placeholder-slate-400 focus:ring-indigo-500'
                        }`}
                        value={feeInputValue}
                        onChange={(event) => {
                          const value = event.target.value;
                          const singerRecord = poolMap.get(singerId);
                          if (!singerRecord) return;
                          const serviceId = selectedService ? String(selectedService.id) : '';
                          const serviceDetails = serviceId ? singerRecord.serviceFees?.[serviceId] : null;
                          const fallbackFee = serviceDetails?.fee != null
                            ? String(serviceDetails.fee)
                            : singerRecord.fee != null ? String(singerRecord.fee) : '';

                          if (!isSelected) {
                            const nextEntries = [
                              ...selectedEntries,
                              {
                                id: singerId,
                                name: singerRecord.name,
                                fee: value === '' ? fallbackFee : value,
                                locked: false
                              }
                            ];
                            updateSelected(nextEntries);
                            return;
                          }

                          const next = selectedEntries.map(entry => (
                            entry.id === singerId ? { ...entry, fee: value } : entry
                          ));
                          updateSelected(next);
                        }}
                        onClick={(event) => event.stopPropagation()}
                        inputMode="decimal"
                      />
                    </div>
                  </label>

                  <div className="ml-auto flex items-center gap-2 flex-shrink-0">
                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        const next = selectedEntries.map(entry => (
                          entry.id === singerId
                            ? { ...entry, locked: !Boolean(entry.locked) }
                            : entry
                        ));
                        if (!isSelected) {
                          handleToggleSinger(singerId);
                          return;
                        }
                        updateSelected(next);
                      }}
                      disabled={!isSelected}
                      className={`inline-flex items-center gap-1 rounded border px-2 py-1 text-xs font-medium focus:outline-none focus:ring-2 ${
                        isSelected
                          ? 'border-white/60 text-white focus:ring-white/40'
                          : 'border-slate-200 text-slate-500 focus:ring-indigo-500 disabled:opacity-60'
                      }`}
                      aria-label={isLocked ? 'Unlock fee' : 'Lock fee'}
                    >
                      {isLocked ? '🔒 Locked' : '🔓 Unlocked'}
                    </button>

                    {canManagePool ? (
                      <button
                        type="button"
                        onClick={(event) => {
                          event.stopPropagation();
                          handleOpenEditSingerModal(singer);
                        }}
                        className={`inline-flex items-center gap-1 rounded border px-2 py-1 text-xs font-semibold focus:outline-none focus:ring-2 ${
                          isSelected ? 'border-white/60 text-white hover:text-indigo-100 focus:ring-white/40' : 'border-indigo-200 text-indigo-600 hover:text-indigo-500 focus:ring-indigo-200'
                        }`}
                      >
                        Edit
                      </button>
                    ) : null}
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="rounded border border-dashed border-slate-300 bg-slate-50 px-3 py-2 text-sm text-slate-500">
            No singers available yet. Add them below.
          </div>
        )}
      </section>

      <div className="rounded border border-slate-200 bg-white p-3 text-sm space-y-2">
        <div className="flex items-center justify-between">
          <span className="font-medium text-slate-600">Singer discount</span>
          {singerDiscountType === 'percent' && singerDiscountValueNumber > 0 ? (
            <span className="text-xs text-slate-500">≈ {toCurrency(singerDiscountValueNumber)}</span>
          ) : null}
        </div>
        <div className="flex gap-1 w-32 sm:w-36">
          {['amount', 'percent'].map(type => (
            <button
              key={type}
              type="button"
              onClick={() => {
                if (type !== singerDiscountType) onChange('pricing_discount_type', type);
              }}
              className={`inline-flex flex-1 items-center justify-center rounded-full border px-2.5 py-1 text-xs font-medium transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${
                type === singerDiscountType ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:border-indigo-200 hover:text-indigo-600'
              }`}
            >
              {type === 'amount' ? 'Amount (£)' : 'Percent (%)'}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          <div className="relative w-32 sm:w-36">
            <span className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-xs text-slate-400">
              {singerDiscountType === 'amount' ? '£' : '%'}
            </span>
            <input
              type="number"
              step="0.01"
              className="w-full rounded border border-slate-300 px-6 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              value={formState.pricing_discount || ''}
              onChange={event => onChange('pricing_discount', event.target.value)}
            />
          </div>
          {singerDiscountType === 'percent' ? (
            <span className="text-xs text-slate-500">≈ {toCurrency(singerDiscountValueNumber)}</span>
          ) : null}
        </div>
      </div>


      <div className="rounded-lg bg-indigo-50 p-3 text-sm text-indigo-700">
        <div className="font-semibold">Quote summary</div>
        <div>{totals.singerCount} singer{totals.singerCount === 1 ? '' : 's'} selected · Base fee {toCurrency(totals.base)}</div>
        <div>Singer fees after discount: {toCurrency(totals.singerNet ?? singerNetValue)}</div>
        <div>Production after discount: {toCurrency(totals.productionNet ?? productionNetValue)}</div>
        <div>Singer discount: -{toCurrency(singerDiscountValueNumber)}</div>
        <div>Production discount: -{toCurrency(productionDiscountValueNumber)}</div>
        <div>Custom fees: {toCurrency(customFeesNumber)}</div>
        <div className="font-semibold text-indigo-900">Total after adjustments: {toCurrency(totalValue)}</div>
      </div>
      </div>

      {showAddSingerModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4">
          <div className="w-full max-w-md rounded-lg bg-white p-6 shadow-xl">
            <div className="flex items-start justify-between">
              <div>
                <h3 className="text-lg font-semibold text-slate-800">Add singer to pool</h3>
                <p className="text-sm text-slate-500">Capture singer details to make them available for future lineups.</p>
              </div>
              <button
                type="button"
                onClick={handleCloseAddSingerModal}
                className="text-slate-400 hover:text-slate-600"
                aria-label="Close add singer modal"
                disabled={addingSinger}
              >
                ✕
              </button>
            </div>
            <div className="mt-4 space-y-4">
              <label className="block text-sm font-medium text-slate-600">
                Name
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  value={newSingerName}
                  onChange={event => setNewSingerName(event.target.value)}
                  placeholder="Singer name"
                  disabled={addingSinger}
                />
              </label>
              <label className="block text-sm font-medium text-slate-600">
                Base fee (£)
                <div className="mt-1 flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-1">
                  <span className="text-xs text-slate-500">£</span>
                  <input
                    type="number"
                    step="0.01"
                    className="w-full border-0 bg-transparent p-0 text-sm focus:outline-none"
                    value={newSingerBaseFee}
                    onChange={event => setNewSingerBaseFee(event.target.value)}
                    placeholder="0.00"
                    disabled={addingSinger}
                  />
                </div>
              </label>
              <div className="space-y-3">
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Service pricing</div>
                {serviceTypes.length ? serviceTypes.map(service => {
                  const serviceId = service.id != null ? String(service.id) : '';
                  if (!serviceId) return null;
                  const config = newSingerServiceFees[serviceId] || { fee: '', defaultIncluded: false };
                  return (
                    <div key={service.id} className="rounded border border-slate-200 bg-slate-50 px-3 py-2 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-sm font-medium text-slate-700">{service.label}</span>
                        <label className="flex items-center gap-1 text-[12px] text-slate-600">
                          <input
                            type="checkbox"
                            className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                            checked={Boolean(config.defaultIncluded)}
                            onChange={event => handleNewSingerServiceFeeChange(serviceId, 'defaultIncluded', event.target.checked)}
                            disabled={addingSinger}
                          />
                          Default lineup
                        </label>
                      </div>
                      <div className="flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-1">
                        <span className="text-xs text-slate-500">£</span>
                        <input
                          type="number"
                          step="0.01"
                          className="w-full border-0 bg-transparent p-0 text-sm focus:outline-none"
                          value={config.fee}
                          onChange={event => handleNewSingerServiceFeeChange(serviceId, 'fee', event.target.value)}
                          placeholder="Defaults to base fee"
                          disabled={addingSinger}
                        />
                      </div>
                    </div>
                  );
                }) : (
                  <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-500">
                    No service templates found. Base fee will be used.
                  </div>
                )}
              </div>
              {addError ? (
                <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-600">{addError}</div>
              ) : null}
            </div>
            <div className="mt-6 flex items-center justify-end gap-3">
              <button
                type="button"
                onClick={handleCloseAddSingerModal}
                className="text-sm font-medium text-slate-600 hover:text-slate-800 disabled:opacity-60"
                disabled={addingSinger}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleAddSingerToPool}
                disabled={confirmDisabled}
                className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
              >
                {addingSinger ? 'Saving…' : 'Add to pool'}
              </button>
            </div>
            {!canManagePool ? (
              <p className="mt-3 text-xs text-slate-500">Pool updates are unavailable in this view.</p>
            ) : null}
          </div>
        </div>
      ) : null}

      {showEditSingerModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl">
            <div className="flex items-start justify-between">
              <div>
                <h3 className="text-lg font-semibold text-slate-800">Edit singer</h3>
                <p className="text-sm text-slate-500">Update singer details and service pricing.</p>
              </div>
              <button
                type="button"
                onClick={handleCloseEditSingerModal}
                className="text-slate-400 hover:text-slate-600"
                aria-label="Close edit singer modal"
                disabled={editingSinger}
              >
                ✕
              </button>
            </div>
            <div className="mt-4 space-y-4">
              <label className="block text-sm font-medium text-slate-600">
                Name
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSingerName}
                  onChange={event => setEditSingerName(event.target.value)}
                  placeholder="Singer name"
                  disabled={editingSinger}
                />
              </label>
              <label className="block text-sm font-medium text-slate-600">
                Base fee (£)
                <div className="mt-1 flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-1">
                  <span className="text-xs text-slate-500">£</span>
                  <input
                    type="number"
                    step="0.01"
                    className="w-full border-0 bg-transparent p-0 text-sm focus:outline-none"
                    value={editSingerBaseFee}
                    onChange={event => setEditSingerBaseFee(event.target.value)}
                    placeholder="0.00"
                    disabled={editingSinger}
                  />
                </div>
              </label>
              <label className="flex items-center gap-2 text-sm text-slate-600">
                <input
                  type="checkbox"
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                  checked={editSingerDefaultIncluded}
                  onChange={event => setEditSingerDefaultIncluded(event.target.checked)}
                  disabled={editingSinger}
                />
                Default lineup when no service is selected
              </label>
              <div className="space-y-3">
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Service pricing</div>
                {serviceTypes.map(service => {
                  const serviceId = service.id != null ? String(service.id) : '';
                  if (!serviceId) return null;
                  const config = editSingerServiceFees[serviceId] || { fee: '', defaultIncluded: false };
                  return (
                    <div key={service.id} className="rounded border border-slate-200 bg-slate-50 px-3 py-2 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-sm font-medium text-slate-700">{service.label}</span>
                        <label className="flex items-center gap-1 text-[12px] text-slate-600">
                          <input
                            type="checkbox"
                            className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                            checked={Boolean(config.defaultIncluded)}
                            onChange={event => handleEditServiceFeeChange(serviceId, 'defaultIncluded', event.target.checked)}
                            disabled={editingSinger}
                          />
                          Default lineup
                        </label>
                      </div>
                      <div className="flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-1">
                        <span className="text-xs text-slate-500">£</span>
                        <input
                          type="number"
                          step="0.01"
                          className="w-full border-0 bg-transparent p-0 text-sm focus:outline-none"
                          value={config.fee ?? ''}
                          onChange={event => handleEditServiceFeeChange(serviceId, 'fee', event.target.value)}
                          placeholder="Defaults to base fee"
                          disabled={editingSinger}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
              {editError ? (
                <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-600">{editError}</div>
              ) : null}
            </div>
            <div className="mt-6 flex items-center justify-between">
              <button
                type="button"
                onClick={handleDeleteSinger}
                disabled={editingSinger}
                className="text-sm font-medium text-red-600 hover:text-red-500 disabled:opacity-60"
              >
                Delete singer
              </button>
              <div className="flex items-center gap-3">
                <button
                  type="button"
                  onClick={handleCloseEditSingerModal}
                  className="text-sm font-medium text-slate-600 hover:text-slate-800 disabled:opacity-60"
                  disabled={editingSinger}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleSaveEditedSinger}
                  disabled={editingSinger || !editSingerName.trim()}
                  className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
                >
                  {editingSinger ? 'Saving…' : 'Save changes'}
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}

function GigInfoPanel({ formState, onChange, businessId, jobsheetId }) {
  const parseJson = (raw) => {
    if (!raw) return { values: {}, include: {} };
    try { const obj = JSON.parse(raw); return obj && typeof obj === 'object' ? obj : { values: {}, include: {} }; } catch (_) { return { values: {}, include: {} }; }
  };
  const initial = parseJson(formState?.gig_info || '');
  const [values, setValues] = useState({ ...(initial.values || {}) });
  const [include, setInclude] = useState({ ...(initial.include || {}) });
  const [gigToasts, setGigToasts] = useState([]);
  const pushGigToast = (text, tone = 'info') => {
    const notice = { id: `gig-toast-${Date.now()}-${Math.random().toString(36).slice(2)}`, text, tone };
    setGigToasts(prev => [...prev, notice]);
    setTimeout(() => setGigToasts(prev => prev.filter(t => t !== notice)), 3000);
  };

  // Helpers for time formatting and call-time computation
  const fmtTime = (input) => {
    if (!input) return '';
    let s = String(input).trim();
    if (!s) return '';
    s = s.replace(/\./g, ':').replace(/\s+/g, '');
    let mer = null;
    const lower = s.toLowerCase();
    if (/(am|pm)$/.test(lower)) {
      mer = lower.slice(-2);
      s = lower.slice(0, -2);
    }
    let h = 0; let m = 0;
    if (/^\d{1,2}:\d{2}$/.test(s)) {
      const parts = s.split(':');
      h = Number(parts[0]);
      m = Number(parts[1]);
    } else if (/^\d{3,4}$/.test(s)) {
      const v = s.padStart(4, '0');
      h = Number(v.slice(0, 2));
      m = Number(v.slice(2));
    } else if (/^\d{1,2}$/.test(s)) {
      h = Number(s);
      m = 0;
    } else {
      return String(input);
    }
    if (Number.isNaN(h) || Number.isNaN(m)) return '';
    if (mer) {
      if (mer === 'pm' && h < 12) h += 12;
      if (mer === 'am' && h === 12) h = 0;
    }
    h = Math.max(0, Math.min(23, h));
    m = Math.max(0, Math.min(59, m));
    const outMer = h >= 12 ? 'pm' : 'am';
    const h12 = (h % 12) === 0 ? 12 : (h % 12);
    const mm = String(m).padStart(2, '0');
    return `${h12}:${mm} ${outMer}`;
  };
  const parseMinutes = (input) => {
    if (!input) return null;
    let s = String(input).trim();
    if (!s) return null;
    s = s.replace(/\./g, ':').replace(/\s+/g, '');
    let mer = null;
    const lower = s.toLowerCase();
    if (/(am|pm)$/.test(lower)) { mer = lower.slice(-2); s = lower.slice(0, -2); }
    let h = 0; let m = 0;
    if (/^\d{1,2}:\d{2}$/.test(s)) { const parts = s.split(':'); h = Number(parts[0]); m = Number(parts[1]); }
    else if (/^\d{3,4}$/.test(s)) { const v = s.padStart(4, '0'); h = Number(v.slice(0,2)); m = Number(v.slice(2)); }
    else if (/^\d{1,2}$/.test(s)) { h = Number(s); m = 0; }
    else { return null; }
    if (Number.isNaN(h) || Number.isNaN(m)) return null;
    if (mer) { if (mer === 'pm' && h < 12) h += 12; if (mer === 'am' && h === 12) h = 0; }
    h = Math.max(0, Math.min(23, h)); m = Math.max(0, Math.min(59, m));
    return h * 60 + m;
  };
  const fmtFromMinutes = (mins) => {
    if (mins == null) return '';
    let v = ((mins % 1440) + 1440) % 1440;
    const h = Math.floor(v / 60); const m = v % 60;
    const outMer = h >= 12 ? 'pm' : 'am';
    const h12 = (h % 12) === 0 ? 12 : (h % 12);
    const mm = String(m).padStart(2, '0');
    return `${h12}:${mm} ${outMer}`;
  };

  // Prefill from jobsheet on first mount for missing fields
  useEffect(() => {
    setValues(prev => ({
      ...prev,
      client_name: prev.client_name ?? (formState.client_name ?? ''),
      event_type: prev.event_type ?? (formState.event_type ?? '')
    }));
    // Default to including schedule and time lines (prefilled once)
    setInclude(prev => ({
      ...prev,
      client_name: prev.client_name ?? Boolean(formState.client_name),
      event_type: prev.event_type ?? Boolean(formState.event_type),
      event_date: prev.event_date ?? true,
      schedule: prev.schedule ?? true,
      event_time: prev.event_time ?? true,
      call_time: prev.call_time ?? true,
      personnel_lineup: prev.personnel_lineup ?? false,
      repertoire: prev.repertoire ?? false,
      compact_spacing: prev.compact_spacing ?? false
    }));
    // Seed default editable lines if absent
    setValues(prev => {
      const next = { ...prev };
      const tStart = fmtTime(formState.event_start);
      const tEnd = fmtTime(formState.event_end);
      const startMins = parseMinutes(formState.event_start);
      if ((next.event_time == null || String(next.event_time).trim() === '') && (tStart || tEnd)) {
        next.event_time = tStart && tEnd
          ? `Event time: ${tStart} – ${tEnd}`
          : (tStart ? `Event time: ${tStart}` : (tEnd ? `Event end: ${tEnd}` : ''));
      }
      if ((next.call_time == null || String(next.call_time).trim() === '') && startMins != null) {
        next.call_time = `Call time: ${fmtFromMinutes(startMins - 75)}`;
      }
      // Prefill personnel lineup if not set
      try {
        if (next.personnel_lineup == null || String(next.personnel_lineup).trim() === '') {
          const arr = Array.isArray(formState.pricing_selected_singers) ? formState.pricing_selected_singers : [];
          const names = arr.map(e => (e && typeof e === 'object' ? (e.name || '') : '')).filter(Boolean);
          if (names.length) {
            next.personnel_lineup = names.join(', ');
            setInclude(prevInc => ({ ...prevInc, personnel_lineup: true }));
          }
        }
      } catch (_) {}
      return next;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Push to form state (autosave handles persistence) without creating an update loop
  const onChangeRef = useRef(onChange);
  useEffect(() => { onChangeRef.current = onChange; }, [onChange]);
  useEffect(() => {
    try { onChangeRef.current?.('gig_info', JSON.stringify({ values, include })); } catch (_) {}
  }, [values, include]);

  const setVal = (key, v) => setValues(prev => ({ ...prev, [key]: v }));
  const setInc = (key, v) => setInclude(prev => ({ ...prev, [key]: v }));
  const [showPreview, setShowPreview] = useState(true);

  // deprecated standalone generate (replaced by single quick action)

  const [shareWorking, setShareWorking] = useState(false);
  const lastPdfPathRef = useRef(null);
  const handleRevealOnly = async () => {
    try {
      setShareWorking(true);
      // Reveal an existing Gig Info PDF only (do not generate)
      let pdfPath = null;
      try {
        const files = await window.api?.listJobFolderFiles?.({ businessId, jobsheetId, extensionPattern: '\\.(pdf)$' });
        const gigInfos = (Array.isArray(files) ? files : []).filter(f => String(f?.name || '').toLowerCase().startsWith('gig info'));
        if (gigInfos.length) {
          pdfPath = gigInfos[0].path; // listJobFolderFiles returns sorted by mtime desc
        }
      } catch (_) {}
      if (!pdfPath) {
        pushGigToast('No Gig Info PDF found to reveal', 'warning');
        return;
      }
      lastPdfPathRef.current = pdfPath;
      const reveal = await window.api?.showItemInFolder?.(pdfPath);
      if (reveal && reveal.ok === false) {
        // Fallback: open the file if reveal failed
        await window.api?.openPath?.(pdfPath);
      }
      pushGigToast('Revealed in Finder', 'success');
    } catch (err) {
      pushGigToast(err?.message || 'Unable to reveal PDF', 'error');
    } finally {
      setShareWorking(false);
    }
  };
  const handleGenerateOnly = async () => {
    try {
      setShareWorking(true);
      pushGigToast('Generating PDF…');
      const res = await window.api?.createGigInfoPdf?.({ businessId, jobsheetId, gigInfo: { values, include } });
      if (!res || res.ok === false || !res.file_path) throw new Error(res?.message || 'Unable to generate PDF');
      lastPdfPathRef.current = res.file_path;
      pushGigToast('Gig Info PDF generated', 'success');
    } catch (err) {
      pushGigToast(err?.message || 'Unable to generate PDF', 'error');
    } finally {
      setShareWorking(false);
    }
  };

  // (deprecated individual actions removed; single Share action defined above)

  // Load presets for dress code and repertoire
  const [dressPresets, setDressPresets] = useState([]);
  const [repPresets, setRepPresets] = useState([]);
  const refreshPresets = useCallback(async () => {
    try {
      const data = await window.api?.getGigInfoPresets?.({ businessId });
      setDressPresets(Array.isArray(data?.dress_codes) ? data.dress_codes : []);
      setRepPresets(Array.isArray(data?.repertoire) ? data.repertoire : []);
    } catch (_) {
      setDressPresets([]); setRepPresets([]);
    }
  }, [businessId]);
  useEffect(() => { refreshPresets(); }, [refreshPresets]);

  // Keep dress_code string in sync with explicit selections to avoid stray items
  useEffect(() => {
    try {
      const presets = Array.isArray(dressPresets) ? new Set(dressPresets) : new Set();
      const selected = Array.isArray(values.dress_code_items)
        ? values.dress_code_items.filter(x => presets.size ? presets.has(x) : true)
        : String(values.dress_code || '')
            .split(/[,•\n]+/)
            .map(s => s.trim())
            .filter(Boolean)
            .filter(x => presets.size ? presets.has(x) : true);
      const joined = selected.join(', ');
      if (joined !== String(values.dress_code || '')) {
        setVal('dress_code', joined);
      }
      if (JSON.stringify(selected) !== JSON.stringify(values.dress_code_items || [])) {
        setVal('dress_code_items', selected);
      }
    } catch (_) {}
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dressPresets]);

  const renderPreview = () => {
    const header = include.event_date !== false ? `Gig info sheet: ${formatDateDisplay(formState.event_date)}` : 'Gig info sheet';
    const blocks = [];
    if (include.client_name && (values.client_name || formState.client_name)) {
      blocks.push({ label: 'Client', value: values.client_name || formState.client_name });
    }
    if (include.event_type && (values.event_type || formState.event_type)) {
      blocks.push({ label: 'Event', value: values.event_type || formState.event_type });
    }
    if (include.venue_block !== false) {
      const lines = [
        values.venue_name || formState.venue_name,
        values.venue_address1 || formState.venue_address1,
        values.venue_address2 || formState.venue_address2,
        values.venue_address3 || formState.venue_address3,
        [values.venue_town || formState.venue_town, values.venue_postcode || formState.venue_postcode].filter(Boolean).join(' ')
      ].filter(Boolean);
      if (lines.length) blocks.push({ label: 'Venue', value: lines.join('\n') });
    }
    // Build editable schedule block from stored values
    const schedulePieces = [];
    const evTime = include.event_time !== false ? String(values.event_time || '').trim() : '';
    const call = include.call_time !== false ? String(values.call_time || '').trim() : '';
    const schedText = include.schedule ? String(values.schedule || '').trim() : '';
    if (evTime) schedulePieces.push(evTime);
    if (call) schedulePieces.push(call);
    if (schedText) schedulePieces.push(schedText);
    if (schedulePieces.length) blocks.push({ label: 'Schedule', value: schedulePieces.join('\n') });
    if (include.personnel_lineup && String(values.personnel_lineup || '').trim()) { blocks.push({ label: 'Personnel', value: String(values.personnel_lineup || '').trim() }); }
    if (include.repertoire && String(values.repertoire || '').trim()) { blocks.push({ label: 'Setlist / Repertoire', value: String(values.repertoire || '').trim() }); }
    if (include.dress_code && values.dress_code) blocks.push({ label: 'Dress code', value: values.dress_code });
    if (include.kit_notes && values.kit_notes) blocks.push({ label: 'Kit', value: values.kit_notes });
    if (include.contacts && (values.contractor_name || values.contractor_phone || values.venue_contact_name || values.venue_contact_phone)) {
      const lines = [];
      if (values.contractor_name || values.contractor_phone) lines.push([values.contractor_name, values.contractor_phone].filter(Boolean).join(' · '));
      if (values.venue_contact_name || values.venue_contact_phone) lines.push([values.venue_contact_name, values.venue_contact_phone].filter(Boolean).join(' · '));
      blocks.push({ label: 'Contacts', value: lines.join('\n') });
    }
    if (include.notes && values.notes) blocks.push({ label: 'Notes', value: values.notes });
    const compact = include.compact_spacing === true;
    return (
      <div className={`rounded border border-slate-200 bg-white ${compact ? 'p-2' : 'p-3'}`}>
        <div className="text-base font-semibold text-slate-800 mb-2">{header}</div>
        <div className={compact ? 'space-y-1' : 'space-y-2'}>
          {blocks.map((b, i) => (
            <div key={i}>
              <div className="text-[11px] uppercase tracking-wide text-slate-500">{b.label}</div>
              <div className="whitespace-pre-wrap text-sm text-slate-800">{b.value}</div>
            </div>
          ))}
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      <div className="rounded border border-slate-200 bg-white p-3">
        <div className="text-sm font-semibold text-slate-700">Header</div>
        <label className="mt-2 flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={include.event_date !== false} onChange={e => setInc('event_date', e.target.checked)} />
          Include event date in header (Gig info sheet: {formatDateDisplay(formState.event_date)})
        </label>
        <label className="mt-2 flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.compact_spacing} onChange={e => setInc('compact_spacing', e.target.checked)} />
          Compact spacing (tighter PDF layout)
        </label>
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-3">
        <div className="text-sm font-semibold text-slate-700">Client</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.client_name} onChange={e => setInc('client_name', e.target.checked)} />
          Include client name
        </label>
        <input
          className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm"
          value={values.client_name || ''}
          onChange={e => setVal('client_name', e.target.value)}
          placeholder="Client name"
        />
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-3">
        <div className="text-sm font-semibold text-slate-700">Event</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.event_type} onChange={e => setInc('event_type', e.target.checked)} />
          Include event type
        </label>
        <input
          className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm"
          value={values.event_type || ''}
          onChange={e => setVal('event_type', e.target.value)}
          placeholder="Event type (e.g., Wedding, Corporate)"
        />
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Venue</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={include.venue_block !== false} onChange={e => setInc('venue_block', e.target.checked)} />
          Include venue block
        </label>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Venue name" value={values.venue_name || ''} onChange={e => setVal('venue_name', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Address line 1" value={values.venue_address1 || ''} onChange={e => setVal('venue_address1', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Address line 2" value={values.venue_address2 || ''} onChange={e => setVal('venue_address2', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Address line 3" value={values.venue_address3 || ''} onChange={e => setVal('venue_address3', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Town/City" value={values.venue_town || ''} onChange={e => setVal('venue_town', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Postcode" value={values.venue_postcode || ''} onChange={e => setVal('venue_postcode', e.target.value)} />
        </div>
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Schedule</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.schedule} onChange={e => setInc('schedule', e.target.checked)} />
          Include schedule
        </label>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={include.event_time !== false} onChange={e => setInc('event_time', e.target.checked)} />
            Include event time
          </label>
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Event time: 6:00 pm – 8:30 pm" value={values.event_time || ''} onChange={e => setVal('event_time', e.target.value)} />
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={include.call_time !== false} onChange={e => setInc('call_time', e.target.checked)} />
            Include call time
          </label>
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Call time: 4:45 pm" value={values.call_time || ''} onChange={e => setVal('call_time', e.target.value)} />
        </div>
        <textarea rows={4} className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder={'- 18:00 Call\n- 19:30 Set 1\n- 20:15 Break\n- 20:45 Set 2'} value={values.schedule || ''} onChange={e => setVal('schedule', e.target.value)} />
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Personnel</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.personnel_lineup} onChange={e => setInc('personnel_lineup', e.target.checked)} />
          Include personnel lineup
        </label>
        <textarea
          rows={3}
          className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm"
          placeholder={'Lead: Alice\nTenor: Bob\nBaritone: Carlos\nBass: Dan'}
          value={values.personnel_lineup || ''}
          onChange={e => setVal('personnel_lineup', e.target.value)}
        />
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Setlist / Repertoire</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.repertoire} onChange={e => setInc('repertoire', e.target.checked)} />
          Include setlist / repertoire
        </label>
        <div className="space-y-2">
          <div className="flex flex-wrap items-center gap-2 min-w-0">
            <select
              className="shrink-0 rounded border border-slate-300 px-2 py-1.5 text-sm w-40"
              value=""
              onChange={e => { const v = e.target.value || ''; if (v) setVal('repertoire', v); }}
            >
              <option value="">Preset…</option>
              {repPresets.map((p, i) => {
                const label = (p.split(/\n/)[0] || p);
                return <option key={i} value={p}>{label.length > 40 ? label.slice(0, 40) + '…' : label}</option>;
              })}
            </select>
            <button
              type="button"
              className="shrink-0 rounded border border-slate-300 px-2.5 py-1.5 text-xs text-slate-600 hover:bg-slate-50 whitespace-nowrap"
              onClick={async () => { try { const v = String(values.repertoire || '').trim(); if (!v) return; await window.api?.saveGigInfoPreset?.({ businessId, kind: 'repertoire', value: v }); refreshPresets(); } catch (_) {} }}
            >
              Save current
            </button>
          </div>
          <textarea
            rows={4}
            className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm"
            placeholder={'- All You Need Is Love\n- Stand By Me\n- Can’t Help Falling In Love'}
            value={values.repertoire || ''}
            onChange={e => setVal('repertoire', e.target.value)}
          />
        </div>
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Dress & Kit</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={!!include.dress_code} onChange={e => setInc('dress_code', e.target.checked)} />
            Include dress code
          </label>
          <div className="space-y-2">
            {/* Checklist of dress code items */}
            <div className="flex flex-wrap gap-3">
              {dressPresets.length ? dressPresets.map((item, idx) => {
                const selectedTokens = Array.isArray(values.dress_code_items)
                  ? values.dress_code_items
                  : String(values.dress_code || '').split(/[,•\n]+/).map(s => s.trim()).filter(Boolean);
                const selectedSet = new Set(selectedTokens);
                const checked = selectedSet.has(item);
                return (
                  <label key={idx} className="inline-flex items-center gap-1.5 text-sm text-slate-700">
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={(e) => {
                        const next = new Set(selectedSet);
                        if (e.target.checked) next.add(item); else next.delete(item);
                        const arr = Array.from(next);
                        setVal('dress_code', arr.join(', '));
                        setVal('dress_code_items', arr);
                      }}
                    />
                    <span>{item}</span>
                    <button
                      type="button"
                      title="Delete"
                      className="ml-1 rounded border border-slate-300 px-1 text-xs text-slate-500 hover:bg-slate-100"
                      onClick={async () => { try { await window.api?.deleteGigInfoPreset?.({ businessId, kind: 'dress_code', value: item }); refreshPresets(); } catch (_) {} }}
                    >
                      ✕
                    </button>
                  </label>
                );
              }) : (
                <div className="text-xs text-slate-500">No dress code items yet.</div>
              )}
            </div>
            {/* Add new & rename */}
            <div className="flex flex-wrap items-center gap-2">
              <input
                className="rounded border border-slate-300 px-2 py-1 text-sm"
                placeholder="Add new item…"
                value={values.__draftDressItem || ''}
                onChange={e => setVal('__draftDressItem', e.target.value)}
              />
              <button
                type="button"
                className="rounded border border-slate-300 px-2.5 py-1 text-xs text-slate-600 hover:bg-slate-50"
                onClick={async () => { try { const v = String(values.__draftDressItem || '').trim(); if (!v) return; await window.api?.saveGigInfoPreset?.({ businessId, kind: 'dress_code', value: v }); setVal('__draftDressItem', ''); refreshPresets(); } catch (_) {} }}
              >
                Add
              </button>
            </div>
          </div>
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={!!include.kit_notes} onChange={e => setInc('kit_notes', e.target.checked)} />
            Include kit notes
          </label>
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Kit notes" value={values.kit_notes || ''} onChange={e => setVal('kit_notes', e.target.value)} />
        </div>
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Contacts</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.contacts} onChange={e => setInc('contacts', e.target.checked)} />
          Include contacts
        </label>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Contractor/MD name" value={values.contractor_name || ''} onChange={e => setVal('contractor_name', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Contractor/MD phone" value={values.contractor_phone || ''} onChange={e => setVal('contractor_phone', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Venue contact name" value={values.venue_contact_name || ''} onChange={e => setVal('venue_contact_name', e.target.value)} />
          <input className="rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Venue contact phone" value={values.venue_contact_phone || ''} onChange={e => setVal('venue_contact_phone', e.target.value)} />
        </div>
      </div>

      <div className="rounded border border-slate-200 bg-white p-3 space-y-2">
        <div className="text-sm font-semibold text-slate-700">Notes</div>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input type="checkbox" checked={!!include.notes} onChange={e => setInc('notes', e.target.checked)} />
          Include notes
        </label>
        <textarea rows={3} className="w-full rounded border border-slate-300 px-3 py-1.5 text-sm" placeholder="Any additional information for the team" value={values.notes || ''} onChange={e => setVal('notes', e.target.value)} />
      </div>

      <div className="flex items-center justify-end gap-2">
        <button type="button" className="rounded border border-slate-300 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50" onClick={() => setShowPreview(v => !v)}>{showPreview ? 'Hide preview' : 'Preview'}</button>
        <button type="button" className="rounded border border-indigo-200 px-3 py-1.5 text-sm font-semibold text-indigo-600 hover:bg-indigo-50 disabled:opacity-60 disabled:cursor-not-allowed" onClick={handleGenerateOnly} disabled={shareWorking}>{shareWorking ? 'Working…' : 'Generate PDF'}</button>
        <button type="button" className="rounded border border-indigo-200 px-3 py-1.5 text-sm font-semibold text-indigo-600 hover:bg-indigo-50 disabled:opacity-60 disabled:cursor-not-allowed" onClick={handleRevealOnly} disabled={shareWorking}>{shareWorking ? 'Working…' : 'Reveal in Finder'}</button>
      </div>
      {showPreview ? renderPreview() : null}
      <ToastOverlay notices={gigToasts} />
    </div>
  );
}
function DocumentsInlinePanel({
  jobsheetId,
  jobsheetStatus,
  documents,
  documentDefinitions,
  loading,
  definitionsLoading,
  error,
  onRefresh,
  onGenerate,
  onExportPdf,
  onToggleLock,
  locksOverride = {},
  onOpenFile,
  onRevealFile,
  onDelete,
  documentFolder,
  businessId,
  lastInvoiceNumber,
  jobsheetSnapshot
}) {
  const INVOICE_GATE_BYPASS_KEY = 'invoiceMaster:bypassInvoiceGate';
  const numericBusinessId = businessId != null ? Number(businessId) : null;
  const numericJobsheetId = jobsheetId != null ? Number(jobsheetId) : null;
  const emailStatusStyles = {
    sent: { label: 'Sent', className: 'bg-green-100 text-green-700 border-green-200' },
    scheduled: { label: 'Scheduled', className: 'bg-indigo-100 text-indigo-700 border-indigo-200' },
    scheduled_error: { label: 'Retrying', className: 'bg-amber-100 text-amber-700 border-amber-200' },
    error: { label: 'Failed', className: 'bg-red-100 text-red-700 border-red-200' }
  };

  const renderEmailStatusPill = (status) => {
    const key = String(status || '').toLowerCase();
    const style = emailStatusStyles[key];
    if (!style) return null;
    return (
      <span className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[11px] font-semibold ${style.className}`}>
        {style.label}
      </span>
    );
  };
  const list = Array.isArray(documents) ? documents : [];
  const excelDocs = list.filter(doc => (doc?.file_path || '').toLowerCase().endsWith('.xlsx'));
  const pdfDocs = list.filter(doc => (doc?.file_path || '').toLowerCase().endsWith('.pdf'));
  const defs = Array.isArray(documentDefinitions) ? documentDefinitions : [];

  // helpers to match PDFs to workbook by base name
  const baseNameNoExt = (fp) => {
    const name = fp ? String(fp).split(/[\\/]+/).pop() : '';
    return name ? name.replace(/\.[^.]+$/, '') : '';
  };
  const normalizeBase = (base) => {
    if (!base) return '';
    let s = String(base);
    // unify dash types
    s = s.replace(/[–—]/g, '-');
    // strip trailing (INV-###)
    s = s.replace(/\s*\(INV[-\s]?\d+\)\s*$/i, '');
    // strip trailing (n) copies
    s = s.replace(/\s*\(\d+\)\s*$/g, '');
    return s.trim();
  };
  const workbookDocsByKey = new Map(
    excelDocs.map(d => [d.definition_key || 'workbook', d])
  );
  const pdfByBase = new Map(
    pdfDocs.map(d => [normalizeBase(baseNameNoExt(d.file_path || '')), d])
  );

  // Dynamic: show all definitions that point to an .xlsx template
  const excelDefs = defs
    .filter(d => (d?.template_path || '').toLowerCase().endsWith('.xlsx'))
    .sort((a, b) => {
      const ao = Number.isFinite(a.sort_order) ? a.sort_order : 0;
      const bo = Number.isFinite(b.sort_order) ? b.sort_order : 0;
      if (ao !== bo) return ao - bo;
      const al = (a.label || a.key || '').toLowerCase();
      const bl = (b.label || b.key || '').toLowerCase();
      return al.localeCompare(bl);
    });

  const excelItems = excelDefs.map(def => {
    const doc = def ? workbookDocsByKey.get(def.key) : null;
    const label = def.label || def.key;
    return { def, doc, label };
  });

  const pdfItems = excelItems.map(({ def, label }) => {
    const wbDoc = def ? workbookDocsByKey.get(def.key) : null;
    const wbBase = normalizeBase(baseNameNoExt(wbDoc?.file_path || ''));
    const pdfDoc = wbDoc ? pdfByBase.get(wbBase) : null;
    return { def, wbDoc, pdfDoc, label };
  });

  const composerStoreKey = jobsheetId != null ? `jobsheet:${jobsheetId}` : 'jobsheet:global';
  const storedComposerState = loadComposerState(composerStoreKey);

  const [menuOpenId, setMenuOpenId] = useState('');
  const menuRef = useRef(null);
  const [overrideNumbers, setOverrideNumbers] = useState({});
  const [defaultNext, setDefaultNext] = useState(null);
  const [localToasts, setLocalToasts] = useState([]);
  const [composerOpen, setComposerOpen] = useState(() => storedComposerState?.open ?? false);
  const [composerMountKey, setComposerMountKey] = useState(0);
  const [composerTo, setComposerTo] = useState(() => storedComposerState?.to ?? '');
  const [composerCc, setComposerCc] = useState(() => storedComposerState?.cc ?? '');
  const [composerBcc, setComposerBcc] = useState(() => storedComposerState?.bcc ?? '');
  const [composerSubject, setComposerSubject] = useState(() => storedComposerState?.subject ?? '');
  const [composerBody, setComposerBody] = useState(() => storedComposerState?.body ?? '');
  const [composerAttachments, setComposerAttachments] = useState(() => {
    const saved = storedComposerState?.attachments;
    return Array.isArray(saved) ? [...saved] : [];
  });
  const [composerTemplateKey, setComposerTemplateKey] = useState('');
  const pendingLockRef = useRef({ workbook: new Set(), pdf: new Set() });
  const [, forcePendingLockTick] = useState(0);
  const [composerSendMode, setComposerSendMode] = useState(() => storedComposerState?.sendMode || 'now');
  const [composerScheduleAt, setComposerScheduleAt] = useState(() => storedComposerState?.scheduleAt || '');
  const [composerIncludeSignature, setComposerIncludeSignature] = useState(() => (
    storedComposerState?.includeSignature !== undefined ? Boolean(storedComposerState.includeSignature) : true
  ));

  const prevComposerKeyRef = useRef(composerStoreKey);
  useEffect(() => {
    if (prevComposerKeyRef.current === composerStoreKey) return;
    prevComposerKeyRef.current = composerStoreKey;
    const restored = loadComposerState(composerStoreKey);
    if (restored) {
      setComposerOpen(Boolean(restored.open));
      setComposerTo(restored.to ?? '');
      setComposerCc(restored.cc ?? '');
      setComposerBcc(restored.bcc ?? '');
      setComposerSubject(restored.subject ?? '');
      setComposerBody(restored.body ?? '');
      setComposerAttachments(Array.isArray(restored.attachments) ? [...restored.attachments] : []);
      setComposerTemplateKey(restored.templateKey ?? '');
      setComposerSendMode(restored.sendMode || 'now');
      setComposerScheduleAt(restored.scheduleAt || '');
      setComposerIncludeSignature(restored.includeSignature !== undefined ? Boolean(restored.includeSignature) : true);
    } else {
      setComposerOpen(false);
      setComposerTo('');
      setComposerCc('');
      setComposerBcc('');
      setComposerSubject('');
      setComposerBody('');
      setComposerAttachments([]);
      setComposerTemplateKey('');
      setComposerSendMode('now');
      setComposerScheduleAt('');
      setComposerIncludeSignature(true);
    }
  }, [composerStoreKey]);

  useEffect(() => {
    if (!composerStoreKey) return;
    if (composerOpen) {
      persistComposerState(composerStoreKey, {
        open: true,
        to: composerTo,
        cc: composerCc,
        bcc: composerBcc,
        subject: composerSubject,
        body: composerBody,
        attachments: Array.isArray(composerAttachments) ? [...composerAttachments] : [],
        templateKey: composerTemplateKey,
        sendMode: composerSendMode,
        scheduleAt: composerScheduleAt,
        includeSignature: composerIncludeSignature
      });
    } else {
      clearComposerState(composerStoreKey);
    }
  }, [composerStoreKey, composerOpen, composerTo, composerCc, composerBcc, composerSubject, composerBody, composerAttachments, composerTemplateKey, composerSendMode, composerScheduleAt, composerIncludeSignature]);

  useEffect(() => () => {
    if (!composerStoreKey || !composerOpen) return;
    persistComposerState(composerStoreKey, {
      open: true,
      to: composerTo,
      cc: composerCc,
      bcc: composerBcc,
      subject: composerSubject,
      body: composerBody,
      attachments: Array.isArray(composerAttachments) ? [...composerAttachments] : [],
      templateKey: composerTemplateKey,
      sendMode: composerSendMode,
      scheduleAt: composerScheduleAt,
      includeSignature: composerIncludeSignature
    });
  }, [composerStoreKey, composerOpen, composerTo, composerCc, composerBcc, composerSubject, composerBody, composerAttachments, composerTemplateKey, composerSendMode, composerScheduleAt, composerIncludeSignature]);

  
  const [emailLog, setEmailLog] = useState([]);
  const [emailLogLoading, setEmailLogLoading] = useState(false);
  // Removed legacy "Other files" listing and import flow

  useEffect(() => {
    const onDoc = (e) => {
      if (!menuOpenId) return;
      if (menuRef.current && menuRef.current.contains(e.target)) return;
      setMenuOpenId('');
    };
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, [menuOpenId]);

  // Bypass gate toggle (persisted)
  const [bypassInvoiceGate, setBypassInvoiceGate] = useState(false);
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(INVOICE_GATE_BYPASS_KEY);
      setBypassInvoiceGate(raw === '1');
    } catch (_) {}
  }, []);
  useEffect(() => {
    try {
      window.localStorage.setItem(INVOICE_GATE_BYPASS_KEY, bypassInvoiceGate ? '1' : '0');
    } catch (_) {}
  }, [bypassInvoiceGate]);

  // (Other files list removed)

  // Load default next number from business settings; update when it changes
  useEffect(() => {
    const val = Number(lastInvoiceNumber);
    if (Number.isInteger(val)) {
      setDefaultNext(val + 1);
    } else {
      setDefaultNext(null);
    }
  }, [lastInvoiceNumber]);

  // (previous simple renderRow removed; panes now render rows inline)
  

  const loadEmailLog = useCallback(async () => {
    try {
      if (!jobsheetId || !window.api?.listEmailLog) { setEmailLog([]); return; }
      setEmailLogLoading(true);
      const rows = await window.api.listEmailLog({ jobsheet_id: jobsheetId, limit: 100 });
      setEmailLog(Array.isArray(rows) ? rows : []);
    } catch (_) {
      setEmailLog([]);
    } finally {
      setEmailLogLoading(false);
    }
  }, [jobsheetId]);

  useEffect(() => { loadEmailLog(); }, [loadEmailLog]);

  useEffect(() => {
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload) return;
      if (payload.businessId != null && businessId != null && Number(payload.businessId) !== Number(businessId)) return;
      if (jobsheetId != null) {
        const payloadJobsheetId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
        if (payloadJobsheetId != null && Number(jobsheetId) !== payloadJobsheetId) return;
      }
      if (payload.type === 'email-log-updated') {
        loadEmailLog();
      }
    });
    return () => unsubscribe?.();
  }, [businessId, jobsheetId, loadEmailLog]);

  // removed booking pack composer

  const openComposer = useCallback((options = {}) => {
    const attachments = Array.isArray(options.attachments) ? options.attachments.filter(Boolean) : [];
    setComposerTo(options.to != null ? options.to : (jobsheetSnapshot?.client_email || ''));
    setComposerCc(options.cc ?? '');
    setComposerBcc(options.bcc ?? '');
    setComposerSubject(options.subject ?? '');
    setComposerBody(options.body ?? '');
    setComposerAttachments(attachments);
    setComposerTemplateKey(options.templateKey ?? '');
    setComposerSendMode(options.sendMode || 'now');
    setComposerScheduleAt(options.scheduleAt || '');
    setComposerIncludeSignature(options.includeSignature !== undefined ? Boolean(options.includeSignature) : true);
    // Force a fresh mount so template selection and content always reflect the latest intent
    setComposerMountKey(key => key + 1);
    setComposerOpen(true);
  }, [jobsheetSnapshot]);

  const queueAutoLock = useCallback((docKey, stage) => {
    if (!docKey) return;
    const targetSet = pendingLockRef.current?.[stage];
    if (!targetSet) return;
    targetSet.add(docKey);
    forcePendingLockTick(tick => tick + 1);
  }, [forcePendingLockTick]);

  const openComposerForPdf = (pdfPath, variant) => {
    const v = String(variant || '').toLowerCase();
    const subject = v
      ? `Invoice (${v}) – ${jobsheetSnapshot?.client_name || 'Client'} – ${formatDateDisplay(jobsheetSnapshot?.event_date)}`
      : `Invoice – ${jobsheetSnapshot?.client_name || 'Client'} – ${formatDateDisplay(jobsheetSnapshot?.event_date)}`;
    openComposer({
      attachments: pdfPath ? [pdfPath] : [],
      subject,
      templateKey: ''
    });
  };
  const pushToast = (text, tone = 'info') => {
    const notice = { id: `toast-${Date.now()}-${Math.random().toString(36).slice(2)}`, text, tone };
    setLocalToasts(prev => [...prev, notice]);
    setTimeout(() => {
      setLocalToasts(prev => prev.filter(t => t !== notice));
    }, 3500);
  };

  const handleDeleteEmail = async (id) => {
    if (!id) return;
    const confirmDelete = window.confirm('Delete this sent email entry?');
    if (!confirmDelete) return;
    try {
      await window.api?.deleteEmailLog?.(id);
      setEmailLog(prev => prev.filter(entry => entry.id !== id));
      pushToast('Email log removed', 'success');
    } catch (err) {
      pushToast(err?.message || 'Unable to delete email log', 'error');
    }
  };

  const handleEditScheduledEmail = async (entry) => {
    if (!entry) return;
    try {
      let attachments = [];
      const rawAttachments = entry.attachments;
      if (typeof rawAttachments === 'string') {
        try {
          const parsed = JSON.parse(rawAttachments);
          if (Array.isArray(parsed)) attachments = parsed.filter(Boolean);
        } catch (_) {
          attachments = rawAttachments ? [rawAttachments] : [];
        }
      } else if (Array.isArray(rawAttachments)) {
        attachments = rawAttachments.filter(Boolean);
      }
      openComposer({
        to: entry.to_address || '',
        cc: entry.cc_address || '',
        bcc: entry.bcc_address || '',
        subject: entry.subject || '',
        body: entry.body || '',
        attachments,
        templateKey: entry.template_key || '',
        sendMode: 'later',
        scheduleAt: entry.sent_at || '',
        includeSignature: composerIncludeSignature
      });
    } catch (err) {
      pushToast(err?.message || 'Unable to load scheduled email', 'error');
    }
  };

  const pdfItemByKey = useMemo(() => {
    const map = new Map();
    (pdfItems || []).forEach(item => {
      const key = item?.def?.key;
      if (key) map.set(key, item);
    });
    return map;
  }, [pdfItems]);

  const emailStatusByAttachment = useMemo(() => {
    const map = new Map();
    (emailLog || []).forEach(entry => {
      let attachments = [];
      try {
        attachments = JSON.parse(entry.attachments || '[]');
      } catch (_) {
        attachments = [];
      }
      const status = String(entry.status || 'sent').toLowerCase();
      attachments
        .map(att => (att != null ? String(att) : ''))
        .filter(Boolean)
        .forEach(path => {
          if (!map.has(path)) {
            map.set(path, { status, entry });
          }
        });
    });
    return map;
  }, [emailLog]);

  const statusKey = normalizeStatus(jobsheetStatus) || 'enquiry';
  const invoiceGateOpen = bypassInvoiceGate || statusKey === 'contracting' || statusKey === 'confirmed' || statusKey === 'completed';

  const documentRows = useMemo(() => {
    const baseRows = excelItems.map(({ def, doc, label }) => {
      const key = def?.key || label;
      const pdfItem = def ? pdfItemByKey.get(def.key) || null : null;
      const pdfDoc = pdfItem?.pdfDoc || null;
      const pdfPath = pdfDoc?.file_path ? String(pdfDoc.file_path) : null;
      const emailInfo = pdfPath ? emailStatusByAttachment.get(pdfPath) || null : null;
      const invoiceVariant = def?.invoice_variant || '';
      const variantLabel = invoiceVariant
        ? invoiceVariant
            .replace(/_/g, ' ')
            .replace(/\b\w/g, (char) => char.toUpperCase())
        : '';
      const isInvoiceDef = def && (def.invoice_variant === 'deposit' || def.invoice_variant === 'balance');
      let mailTemplateKey = '';
      if (def?.key === 'quote') {
        mailTemplateKey = 'quote';
      } else if (def?.key === 'invoice_balance') {
        mailTemplateKey = 'invoice_balance';
      }
      const suppressEmail = key === 'client_data' || BOOKING_PACK_DEFINITION_KEYS.has(def?.key);
      return {
        key,
        def,
        label,
        invoiceVariant,
        variantLabel,
        isInvoiceDef,
        gateOk: isInvoiceDef ? invoiceGateOpen : true,
        workbookDoc: doc || null,
        pdfDoc,
        pdfItem,
        emailInfo,
        workbookGenerated: Boolean(doc?.file_path),
        pdfExported: Boolean(pdfDoc?.file_path),
        pdfPath,
        mailTemplateKey,
        suppressEmail,
        mailScheduledAt: (emailInfo?.status && String(emailInfo.status).toLowerCase() === 'scheduled' && emailInfo.entry?.sent_at) ? emailInfo.entry.sent_at : null
      };
    });

    const bookingPackDocs = [];
    const orderedDocs = [];

    baseRows.forEach(row => {
      if (BOOKING_PACK_DEFINITION_KEYS.has(row.def?.key)) {
        bookingPackDocs.push(row);
      } else {
        orderedDocs.push(row);
      }
    });

    const result = orderedDocs.map(doc => ({ type: 'doc', doc }));
    if (bookingPackDocs.length) {
      // Ensure consistent order inside the booking pack: Booking schedule → T&Cs → Deposit invoice
      const packOrder = new Map([
        ['booking_schedule', 0],
        ['t_cs', 1],
        ['invoice_deposit', 2]
      ]);
      bookingPackDocs.sort((a, b) => {
        const ak = a?.def?.key || '';
        const bk = b?.def?.key || '';
        const ai = packOrder.has(ak) ? packOrder.get(ak) : 999;
        const bi = packOrder.has(bk) ? packOrder.get(bk) : 999;
        if (ai !== bi) return ai - bi;
        return (a.label || '').localeCompare(b.label || '', 'en', { sensitivity: 'base' });
      });
      const groupEntry = {
        type: 'group',
        key: 'booking_pack',
        label: 'Booking pack',
        templateKey: 'booking_pack',
        docs: bookingPackDocs,
        attachments: bookingPackDocs.map(doc => doc.pdfPath).filter(Boolean)
      };
      const quoteIndex = result.findIndex(item => item.type === 'doc' && item.doc?.def?.key === 'quote');
      if (quoteIndex >= 0) {
        result.splice(quoteIndex + 1, 0, groupEntry);
      } else {
        result.push(groupEntry);
      }
    }

    return result;
  }, [excelItems, pdfItemByKey, emailStatusByAttachment, invoiceGateOpen]);

  useEffect(() => {
    const pending = pendingLockRef.current;
    if (!pending) return;
    let consumed = false;
    documentRows.forEach(item => {
      const row = item && item.type === 'doc' ? item.doc : null;
      if (!row) return;
      const key = row.def?.key;
      if (!key) return;
      if (pending.workbook?.has(key) && row.workbookGenerated && row.workbookDoc && !row.workbookDoc.is_locked) {
        pending.workbook.delete(key);
        consumed = true;
        try {
          onToggleLock?.(row.workbookDoc);
        } catch (err) {
          console.warn('Auto-lock workbook failed', err);
        }
      }
      if (pending.pdf?.has(key) && row.pdfExported && row.pdfDoc && !row.pdfDoc.is_locked) {
        pending.pdf.delete(key);
        consumed = true;
        try {
          onToggleLock?.(row.pdfDoc);
        } catch (err) {
          console.warn('Auto-lock PDF failed', err);
        }
      }
    });
    if (consumed) {
      forcePendingLockTick(tick => tick + 1);
    }
  }, [documentRows, onToggleLock, forcePendingLockTick]);

  const renderActionPill = ({ label, onClick, disabled, tone = 'slate', key: keyProp, variant = 'outline' }) => {
    const base = 'inline-flex items-center rounded-full border px-2.5 py-0.5 text-xs font-medium transition focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-1';
    let toneClass = '';
    if (variant === 'solid' && tone === 'indigo') {
      toneClass = 'bg-indigo-600 border-indigo-600 text-white hover:bg-indigo-500';
    } else if (tone === 'indigo') {
      toneClass = 'border-indigo-200 text-indigo-600 hover:bg-indigo-50';
    } else if (tone === 'danger') {
      toneClass = 'border-red-200 text-red-600 hover:bg-red-50';
    } else {
      toneClass = 'border-slate-300 text-slate-600 hover:bg-slate-100';
    }
    return (
      <button
        key={keyProp}
        type="button"
        onClick={onClick}
        disabled={disabled}
        className={`${base} ${toneClass} disabled:cursor-not-allowed disabled:opacity-50`}
      >
        {label}
      </button>
    );
  };

  const readyIcon = (label) => (
    <span
      className="inline-flex h-9 w-9 items-center justify-center rounded-full border border-green-200 bg-green-50 text-lg text-green-600"
      title={label}
      aria-label={label}
    >
      ✓
    </span>
  );

  const scheduleBalanceEmail = useCallback(async (pdfPath) => {
    try {
      if (!jobsheetSnapshot || !numericBusinessId || !numericJobsheetId) return;
      const clientEmail = (jobsheetSnapshot.client_email || '').trim();
      if (!clientEmail) return;

      const reminderDate = jobsheetSnapshot.balance_reminder_date || jobsheetSnapshot.balance_due_date;
      if (!reminderDate) return;
      const sendAt = new Date(`${reminderDate}T09:00:00`);
      if (Number.isNaN(sendAt.valueOf())) return;
      if (sendAt.getTime() < Date.now() + 60 * 1000) {
        // Too close or past; skip auto scheduling
        return;
      }
      openComposer({
        to: clientEmail,
        cc: '',
        bcc: '',
        attachments: pdfPath ? [pdfPath] : [],
        templateKey: 'invoice_balance',
        sendMode: 'later',
        scheduleAt: sendAt.toISOString(),
        includeSignature: composerIncludeSignature
      });
    } catch (err) {
      console.warn('Auto schedule balance email failed', err);
      pushToast(err?.message || 'Unable to prepare balance invoice email', 'error');
    }
  }, [numericBusinessId, numericJobsheetId, jobsheetSnapshot, openComposer, composerIncludeSignature]);

  const renderDocumentRow = (row, { nested = false } = {}) => {
    if (!row) return null;
    const docKey = row.def?.key;
    const workbookDoc = row.workbookDoc;
    const pdfDoc = row.pdfDoc;
    const pdfItem = row.pdfItem;
    const workbookReady = row.workbookGenerated;
    const pdfReady = row.pdfExported;
    const workbookLocked = Boolean(workbookDoc?.is_locked);
    const pdfLocked = Boolean(pdfDoc?.is_locked);
    const emailInfo = row.emailInfo;
    const emailStatusKey = String(emailInfo?.status || '').toLowerCase();
    const mailReady = emailStatusKey ? !['error', 'scheduled_error'].includes(emailStatusKey) : false;
    const mailHasTemplate = Boolean(row.mailTemplateKey);
    const emailEntry = emailInfo?.entry;
    const emailWhen = emailEntry?.sent_at ? formatTimestampDisplay(emailEntry.sent_at) : '';
    const emailBadge = emailInfo ? renderEmailStatusPill(emailStatusKey) : null;
    const emailFallbackLabel = row.pdfExported ? 'No emails' : 'PDF not ready';
    const scheduleDateDisplay = row.mailScheduledAt ? formatTimestampDisplay(row.mailScheduledAt) : '';
    const pdfVariantRequiresNumber = row.def && (row.def.invoice_variant === 'deposit' || row.def.invoice_variant === 'balance');

    const generateDisabled = !jobsheetId || !row.def || !row.def.template_path || definitionsLoading || workbookLocked || !row.gateOk;
    const exportDisabled = !pdfItem || pdfLocked || !row.gateOk || !workbookReady || definitionsLoading;
    const mailDisabled = !mailHasTemplate || !pdfReady;

    const handleWorkbookPrimaryClick = () => {
      if (workbookReady) {
        if (workbookDoc?.file_path) onOpenFile?.(workbookDoc.file_path);
        return;
      }
      if (generateDisabled || !row.def) return;
      if (docKey) queueAutoLock(docKey, 'workbook');
      handleGenerate(row.def.key);
    };

    const handlePdfPrimaryClick = () => {
      if (pdfReady) {
        if (pdfDoc?.file_path) onOpenFile?.(pdfDoc.file_path);
        return;
      }
      if (exportDisabled || !pdfItem) return;
      if (docKey) queueAutoLock(docKey, 'pdf');
      handleExportForDef(pdfItem);
    };

    const handleMailPrimaryClick = async () => {
      if (mailReady || mailDisabled) return;
      const key = row.mailTemplateKey || '';
      openComposer({ templateKey: key, attachments: row.pdfPath && pdfReady ? [row.pdfPath] : [] });
    };

  const lockToggle = (doc, locked, label, key) => (
    <button
      key={key}
      type="button"
      className={`flex h-9 w-9 items-center justify-center rounded border border-slate-300 text-base ${!doc?.document_id ? 'cursor-not-allowed opacity-40' : 'hover:bg-slate-100'}`}
      onClick={() => doc && onToggleLock?.(doc)}
      disabled={!doc?.document_id}
      title={locked ? `Unlock ${label}` : `Lock ${label}`}
    >
      <span aria-hidden>{locked ? '🔒' : '🔓'}</span>
      <span className="sr-only">{locked ? `Unlock ${label}` : `Lock ${label}`}</span>
    </button>
  );

    const workbookRow = (
      <div key="row-workbook" className="flex flex-wrap items-center gap-2 sm:flex-nowrap">
        <span className="w-12 text-xs font-semibold uppercase tracking-wide text-slate-500">XLSX</span>
        {workbookReady ? readyIcon('Workbook ready') : renderActionPill({
          key: `${row.key}-generate`,
          label: 'Generate',
          onClick: handleWorkbookPrimaryClick,
          disabled: generateDisabled,
          tone: 'indigo'
        })}
        {workbookReady ? lockToggle(workbookDoc, workbookLocked, 'Workbook', `${row.key}-workbook-lock`) : null}
        <IconButton
          label="Open workbook"
          onClick={() => onOpenFile?.(workbookDoc?.file_path)}
          disabled={!workbookDoc?.file_path}
          size="md"
          className="border-slate-200 text-slate-600 hover:bg-slate-50"
        >
          <OpenIcon className="h-4 w-4" />
        </IconButton>
        <IconButton
          label="Reveal workbook"
          onClick={() => onRevealFile?.(workbookDoc?.file_path)}
          disabled={!workbookDoc?.file_path}
          size="md"
          className="border-slate-200 text-slate-600 hover:bg-slate-50"
        >
          <RevealIcon className="h-4 w-4" />
        </IconButton>
        {onDelete ? (
          <IconButton
            label="Delete workbook"
            onClick={() => onDelete?.(workbookDoc)}
            disabled={workbookDoc?.document_id == null}
            size="md"
            className="border-red-200 text-red-600 hover:bg-red-50"
          >
            <DeleteIcon className="h-4 w-4" />
          </IconButton>
        ) : null}
      </div>
    );

    const pdfChildren = [
      <span key="label" className="w-12 text-xs font-semibold uppercase tracking-wide text-slate-500">PDF</span>,
      pdfReady ? (
        <span key="tick" className="inline-flex">{readyIcon('PDF ready')}</span>
      ) : renderActionPill({
        key: 'pdf-export',
        label: 'Export',
        onClick: handlePdfPrimaryClick,
        disabled: exportDisabled,
        tone: 'indigo'
      }),
      pdfReady ? lockToggle(pdfDoc, pdfLocked, 'PDF', `${row.key}-pdf-lock`) : null,
      <IconButton
        key="open"
        label="Open PDF"
        onClick={() => onOpenFile?.(pdfDoc?.file_path)}
        disabled={!pdfDoc?.file_path}
        size="md"
        className="border-slate-200 text-slate-600 hover:bg-slate-50"
      >
        <OpenIcon className="h-4 w-4" />
      </IconButton>,
      <IconButton
        key="reveal"
        label="Reveal PDF"
        onClick={() => onRevealFile?.(pdfDoc?.file_path)}
        disabled={!pdfDoc?.file_path}
        size="md"
        className="border-slate-200 text-slate-600 hover:bg-slate-50"
      >
        <RevealIcon className="h-4 w-4" />
      </IconButton>,
      onDelete ? (
        <IconButton
          key="delete"
          label="Delete PDF"
          onClick={() => onDelete?.(pdfDoc)}
          disabled={pdfDoc?.document_id == null}
          size="md"
          className="border-red-200 text-red-600 hover:bg-red-50"
        >
          <DeleteIcon className="h-4 w-4" />
        </IconButton>
      ) : null
    ].filter(Boolean);

    if (!pdfReady && pdfVariantRequiresNumber) {
      pdfChildren.push(
        <label key="invoice-number" className="ml-2 flex items-center gap-1 text-[11px] text-slate-500">
          <span>Invoice #</span>
          <input
            type="number"
            min={1}
            value={overrideNumbers[row.def.key] ?? ''}
            onChange={(e) => setOverrideNumbers(prev => ({ ...prev, [row.def.key]: e.target.value }))}
            placeholder={defaultNext != null ? String(defaultNext) : 'INV #'}
            className="w-24 rounded border border-slate-300 px-2 py-1"
          />
        </label>
      );
    }

    const pdfRow = (
      <div key="row-pdf" className="flex flex-wrap items-center gap-2 sm:flex-nowrap">
        {pdfChildren}
      </div>
    );

    const isBalanceInvoice = row.def && (row.def.invoice_variant || '').toLowerCase() === 'balance';
    const emailControls = [];
    const emailRow = row.suppressEmail ? null : (() => {
      if (mailReady) {
        const statusLabel = emailStatusKey === 'scheduled' ? 'Email scheduled' : 'Email sent';
        emailControls.push(<span key="tick" className="inline-flex">{readyIcon(statusLabel)}</span>);
      }
      if (isBalanceInvoice) {
        const scheduleLabel = emailStatusKey === 'scheduled' ? 'Reschedule' : 'Schedule';
        emailControls.push(renderActionPill({
          key: 'balance-schedule',
          label: scheduleLabel,
          onClick: () => scheduleBalanceEmail(pdfDoc?.file_path || ''),
          disabled: !pdfReady || !pdfDoc?.file_path,
          tone: 'indigo',
          variant: 'solid'
        }));
        if (scheduleDateDisplay) {
          emailControls.push(<span key="scheduled-for" className="text-xs text-slate-500">Scheduled for {scheduleDateDisplay}</span>);
        }
      } else if (!mailReady) {
        emailControls.push(renderActionPill({
          key: 'email-send',
          label: 'Send',
          onClick: handleMailPrimaryClick,
          disabled: mailDisabled,
          tone: 'indigo',
          variant: 'solid'
        }));
        emailControls.push(<span key="fallback" className="text-xs text-slate-500">{emailFallbackLabel}</span>);
      }
      if (emailBadge) {
        emailControls.push(<span key="badge" className="flex items-center">{emailBadge}</span>);
      }
      if (emailWhen) {
        emailControls.push(<span key="when" className="text-xs text-slate-500">{emailWhen}</span>);
      }
      return (
        <div key="row-email" className="flex flex-wrap items-center gap-2">
          <span className="w-12 text-xs font-semibold uppercase tracking-wide text-slate-500">Email</span>
          {emailControls}
        </div>
      );
    })();

    let toneKey = row.def?.doc_type ? String(row.def.doc_type).toLowerCase() : 'default';
    if (row.def?.key === 'client_data') toneKey = 'client_data';
    const tone = DOCUMENT_CARD_TONES[toneKey] || DOCUMENT_CARD_TONES.default;

    return (
      <div
        key={row.key}
        className={`rounded-xl border ${tone.outerBorder} p-3 shadow-sm`}
        style={{ background: `linear-gradient(180deg, rgba(255,255,255,0.9) 0%, ${tone.outerBg} 100%)`, boxShadow: '0 10px 30px rgba(15, 23, 42, 0.08)' }}
      >
        <div className="flex items-center justify-between border-b border-slate-200 pb-2">
          <div className="text-[15px] font-bold tracking-tight text-indigo-800" title={row.label}>{row.label}</div>
        </div>
        <div className="mt-3 space-y-3 text-xs text-slate-600">
          <div className={`rounded border ${tone.innerBorder} bg-white p-2 shadow-sm`}>
            <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:gap-4">
              {workbookRow}
              <div className="hidden sm:block" aria-hidden>
                <div className="h-9 w-px bg-slate-200" />
              </div>
              {pdfRow}
            </div>
          </div>
          {emailRow ? (
            <div className={`rounded border ${tone.innerBorder} bg-white p-2 shadow-sm`}>
              {emailRow}
            </div>
          ) : null}
        </div>
      </div>
    );
  };

  const renderBookingPackGroup = (group) => {
    if (!group) return null;
    const docs = Array.isArray(group.docs) ? group.docs : [];
    const attachments = Array.isArray(group.attachments) ? group.attachments.filter(Boolean) : [];
    const allPdfReady = docs.every(doc => doc.pdfExported);
    const composeDisabled = !allPdfReady || !attachments.length;

    let aggregatedInfo = null;
    attachments.forEach(path => {
      const info = emailStatusByAttachment.get(path) || null;
      if (!info) return;
      if (!aggregatedInfo) {
        aggregatedInfo = info;
        return;
      }
      const currentDate = aggregatedInfo.entry?.sent_at ? new Date(aggregatedInfo.entry.sent_at).valueOf() : 0;
      const nextDate = info.entry?.sent_at ? new Date(info.entry.sent_at).valueOf() : 0;
      if (nextDate > currentDate) {
        aggregatedInfo = info;
      }
    });

    const statusKey = aggregatedInfo?.status ? String(aggregatedInfo.status).toLowerCase() : '';
    const mailReady = statusKey ? !['error', 'scheduled_error'].includes(statusKey) : false;
    const mailBadge = statusKey ? renderEmailStatusPill(statusKey) : null;
    const mailWhen = aggregatedInfo?.entry?.sent_at ? formatTimestampDisplay(aggregatedInfo.entry.sent_at) : '';
    const fallbackLabel = allPdfReady ? 'No emails' : 'PDFs not ready';
    const scheduledFor = statusKey === 'scheduled' && aggregatedInfo?.entry?.sent_at
      ? formatTimestampDisplay(aggregatedInfo.entry.sent_at)
      : '';

    const emailControls = [];
    if (mailReady) {
      emailControls.push(<span key="tick" className="inline-flex">{readyIcon('Booking pack email sent')}</span>);
    } else {
      emailControls.push(renderActionPill({
        key: 'booking-pack-send',
        label: 'Send',
        onClick: () => openComposer({ templateKey: group.templateKey, attachments, includeSignature: composerIncludeSignature }),
        disabled: composeDisabled,
        tone: 'indigo',
        variant: 'solid'
      }));
    }
    if (!mailReady) {
      emailControls.push(<span key="fallback" className="text-xs text-slate-500">{fallbackLabel}</span>);
    }
    if (scheduledFor) {
      emailControls.push(<span key="scheduled-for" className="text-xs text-slate-500">Scheduled for {scheduledFor}</span>);
    }
    if (mailBadge) {
      emailControls.push(<span key="badge" className="flex items-center">{mailBadge}</span>);
    }
    if (mailWhen) {
      emailControls.push(<span key="when" className="text-xs text-slate-500">{mailWhen}</span>);
    }

    return (
      <div
        key={group.key || 'booking-pack'}
        className="rounded-xl border border-indigo-200 p-3 shadow-sm"
        style={{ background: 'linear-gradient(180deg, rgba(255,255,255,0.92) 0%, rgba(224,231,255,0.85) 100%)', boxShadow: '0 10px 30px rgba(15, 23, 42, 0.08)' }}
      >
        <div className="flex items-center justify-between border-b border-slate-200 pb-2">
          <span className="text-[15px] font-bold tracking-tight text-indigo-800">{group.label}</span>
        </div>
        <div className="mt-3 space-y-3 text-xs text-slate-600">
          <div className="rounded border border-indigo-200 bg-white p-2 shadow-sm">
            <div className="flex flex-wrap items-center gap-2">
              <span className="w-12 text-xs font-semibold uppercase tracking-wide text-slate-500">Email</span>
              {emailControls}
            </div>
          </div>
          <div className="space-y-3">
            {docs.map(doc => renderDocumentRow(doc, { nested: true }))}
          </div>
        </div>
      </div>
    );
  };

  const handleGenerate = (key) => onGenerate?.(key);
  const canGenerateAll = Boolean(
    jobsheetId && excelItems.some(i => {
      const isInvoiceDef = i.def && (i.def.invoice_variant === 'deposit' || i.def.invoice_variant === 'balance');
      const gateOk = isInvoiceDef ? invoiceGateOpen : true;
      return gateOk && i.def && i.def.template_path && !(i.doc && i.doc.file_path) && !(i.doc && i.doc.is_locked);
    }) && !definitionsLoading
  );

  const handleExportForDef = async (item) => {
    if (!item) return;
    const { def, wbDoc, pdfDoc } = item;
    const exported = Boolean(pdfDoc && pdfDoc.file_path);
    if (exported || (pdfDoc && pdfDoc.is_locked)) return;
    const isInvoiceDef = def && (def.invoice_variant === 'deposit' || def.invoice_variant === 'balance');
    if (isInvoiceDef && !invoiceGateOpen) {
      error && console.warn('Invoice export gated: move job to Contracting or Confirmed');
      return;
    }
    if (!wbDoc || !wbDoc.file_path) {
      const proceed = window.confirm('No workbook found for this document. Generate it first?');
      if (!proceed) return;
      await onGenerate?.(def.key);
      return;
    }
    const requested = isInvoiceDef ? Number(overrideNumbers[def.key]) : null;
    const opts = isInvoiceDef && Number.isInteger(requested) && requested > 0 ? { requestedNumber: requested } : {};
    const res = await onExportPdf?.(wbDoc, opts);
    if (res && res.ok === false && /exists/i.test(res.message || '')) {
      window.alert(res.message || 'Invoice number conflict. Choose another number.');
    }

  };

  const handleExportAll = async () => {
    // ensure all workbooks exist first
    const missing = pdfItems.filter(i => !i.wbDoc || !i.wbDoc.file_path);
    if (missing.length) {
      const ok = window.confirm('Some PDFs need a workbook. Generate missing workbooks first?');
      if (!ok) return;
      for (const i of missing) { // sequential
        // eslint-disable-next-line no-await-in-loop
        await onGenerate?.(i.def.key);
      }
    }
    // Only export PDFs that are not already exported
    for (const i of pdfItems) {
      const alreadyExported = Boolean(i.pdfDoc && i.pdfDoc.file_path);
      if (i.wbDoc && i.wbDoc.file_path && !alreadyExported) {
        // eslint-disable-next-line no-await-in-loop
        await onExportPdf?.(i.wbDoc);
      }
    }
  };

  return (
    <div className="space-y-4">
      {/* Gate toggle */}
      <div className="flex items-center justify-between">
        <div className="text-sm font-semibold text-slate-700">Documents</div>
        <div className="flex items-center gap-2">
          <label className="inline-flex items-center gap-2 text-xs text-slate-600">
            <input
              type="checkbox"
              checked={bypassInvoiceGate}
              onChange={e => setBypassInvoiceGate(e.target.checked)}
            />
            <span>Bypass invoice export gate</span>
          </label>
          
        </div>
      </div>

      {error ? <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-600">{error}</div> : null}
      <ToastOverlay notices={localToasts} />
      <MailComposer
        key={composerMountKey}
        open={composerOpen}
        onClose={() => setComposerOpen(false)}
        businessId={businessId}
        jobsheetId={jobsheetId}
        initialTo={composerTo}
        initialCc={composerCc}
        initialBcc={composerBcc}
        initialSubject={composerSubject}
        initialBody={composerBody}
        initialAttachments={composerAttachments}
        initialTemplateKey={composerTemplateKey}
        onTemplateChange={setComposerTemplateKey}
        initialSendMode={composerSendMode}
        initialScheduleAt={composerScheduleAt}
        onSendModeChange={setComposerSendMode}
        onScheduleChange={setComposerScheduleAt}
        initialIncludeSignature={composerIncludeSignature}
        onIncludeSignatureChange={setComposerIncludeSignature}
        onSent={(result) => {
          setComposerOpen(false);
          const mode = result?.mode === 'later' ? 'scheduled' : 'sent';
          pushToast(mode === 'scheduled' ? 'Email scheduled' : 'Email sent', 'success');
          loadEmailLog();
        }}
      />
      <div className="space-y-4">
        <div className="rounded border border-slate-200 bg-white">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 px-3 py-2">
            <div className="text-sm font-semibold text-slate-700">Documents</div>
            <div className="flex flex-wrap items-center gap-1.5 text-xs">
              {renderActionPill({ key: 'documents-refresh', label: 'Refresh', onClick: onRefresh })}
              {renderActionPill({
                key: 'documents-generate-missing',
                label: 'Generate missing',
                onClick: () => excelItems.forEach(i => (!i.doc?.file_path && !i.doc?.is_locked && i.def?.template_path) && handleGenerate(i.def.key)),
                disabled: !canGenerateAll,
                tone: 'indigo'
              })}
              {renderActionPill({ key: 'documents-export-all', label: 'Export all PDFs', onClick: handleExportAll, tone: 'indigo' })}
              {renderActionPill({
                key: 'documents-new-email',
                label: 'New email',
                onClick: () => {
                  openComposer({
                    templateKey: '',
                    attachments: []
                  });
                }
              })}
            </div>
          </div>
          <div className="space-y-3">
            {documentRows.length ? documentRows.map(item => {
                  if (item.type === 'group') {
                    return renderBookingPackGroup(item);
                  }
                  if (item.type === 'doc') {
                    return renderDocumentRow(item.doc);
                  }
                  return null;
                }) : (
                  <div className="rounded border border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-500">
                    No document definitions configured for this business.
                  </div>
                )}
          </div>
        </div>

        {/* Email history */}
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <div className="text-sm font-semibold text-slate-700">Email history</div>
            {renderActionPill({ key: 'email-log-refresh', label: 'Refresh', onClick: loadEmailLog })}
          </div>
          <div className="rounded border border-slate-200 bg-white p-2 space-y-1">
                {emailLogLoading ? (
                  <div className="px-2 py-2 text-sm text-slate-500">Loading sent emails…</div>
                ) : (emailLog.length === 0 ? (
                  <div className="px-2 py-2 text-sm text-slate-500">No emails sent yet.</div>
                ) : emailLog.map(entry => {
                  let attList = [];
                  try { attList = JSON.parse(entry.attachments || '[]'); } catch (_) { attList = []; }
                  const attLabel = attList.length ? `${attList.length} attachment${attList.length === 1 ? '' : 's'}` : 'No attachments';
                  const status = String(entry.status || 'sent').toLowerCase();
                  const statusPill = renderEmailStatusPill(status);
                  const when = formatTimestampDisplay(entry.sent_at);
                  const baseInfo = (() => {
                    if (status === 'scheduled') return `Scheduled for ${when || '(pending time)'}`;
                    if (status === 'scheduled_error') return `Retrying soon · planned for ${when || '(pending time)'}`;
                    if (status === 'error') return `Failed at ${when || '(unknown time)'}`;
                    return when || '(unknown time)';
                  })();
                  const detail = `${baseInfo} · to ${entry.to_address || '(unknown)'}${attLabel ? ` · ${attLabel}` : ''}`;
                  const isScheduled = status === 'scheduled';
                  return (
                    <div key={entry.id} className="flex items-center justify-between rounded px-2 py-2">
                      <div className="min-w-0 pr-2">
                        <div className="flex items-center gap-2 text-sm font-medium truncate text-slate-700">
                          <span className="truncate" title={entry.subject || '(no subject)'}>{entry.subject || '(no subject)'}</span>
                          {statusPill}
                        </div>
                        <div className="text-xs text-slate-500 truncate" title={detail}>{detail}</div>
                      </div>
                      <div className="flex items-center gap-2">
                        {isScheduled ? (
                          <button
                            type="button"
                            className="rounded border border-indigo-300 px-2 py-0.5 text-xs text-indigo-600 hover:bg-indigo-50"
                            onClick={() => handleEditScheduledEmail(entry)}
                          >Edit</button>
                        ) : null}
                        <button
                          type="button"
                          className="rounded border border-red-300 px-2 py-0.5 text-xs text-red-600 hover:bg-red-50"
                          onClick={() => handleDeleteEmail(entry.id)}
                        >Delete</button>
                      </div>
                    </div>
                  );
                }))}
          </div>
        </div>
      </div>
      <ToastOverlay notices={localToasts} />
    </div>
  );
}

function ProductionPanel({ formState, onChange, totals }) {
  const productionItems = useMemo(
    () => normalizeProductionItems(formState.pricing_production_items),
    [formState.pricing_production_items]
  );

  const productionTotalValue = useMemo(
    () => calculateProductionTotal(productionItems),
    [productionItems]
  );

  const productionDiscountType = formState.pricing_production_discount_type || 'amount';
  const productionDiscountValueNumber = useMemo(
    () => calculateDiscountValue({
      type: productionDiscountType,
      value: formState.pricing_production_discount,
      subtotal: productionTotalValue
    }),
    [productionDiscountType, formState.pricing_production_discount, productionTotalValue]
  );

  useEffect(() => {
    const nextSubtotalString = productionItems.length ? productionTotalValue.toFixed(2) : '';
    const current = formState.pricing_production_subtotal ?? '';
    if (nextSubtotalString !== current) {
      onChange('pricing_production_subtotal', nextSubtotalString);
    }
  }, [productionItems, productionTotalValue, formState.pricing_production_subtotal, onChange]);

  useEffect(() => {
    const netValue = Math.max(productionTotalValue - productionDiscountValueNumber, 0);
    const hasValues = productionTotalValue > 0 || productionDiscountValueNumber > 0;
    const nextNetString = hasValues ? netValue.toFixed(2) : '';
    const current = formState.pricing_production_total ?? '';
    if (nextNetString !== current) {
      onChange('pricing_production_total', nextNetString);
    }
  }, [productionTotalValue, productionDiscountValueNumber, formState.pricing_production_total, onChange]);

  useEffect(() => {
    const nextDiscountString = productionDiscountValueNumber > 0 ? productionDiscountValueNumber.toFixed(2) : '';
    const current = formState.pricing_production_discount_value ?? '';
    if (nextDiscountString !== current) {
      onChange('pricing_production_discount_value', nextDiscountString);
    }
  }, [productionDiscountValueNumber, formState.pricing_production_discount_value, onChange]);

  const handleAddProductionItem = useCallback(() => {
    const newItem = {
      id: `production-${Date.now()}`,
      name: '',
      description: '',
      cost: '',
      markup: '20',
      notes: ''
    };
    onChange('pricing_production_items', normalizeProductionItems([...productionItems, newItem]));
  }, [productionItems, onChange]);

  const handleProductionChange = useCallback((id, field, value) => {
    const next = productionItems.map(item => (item.id === id ? { ...item, [field]: value } : item));
    onChange('pricing_production_items', normalizeProductionItems(next));
  }, [productionItems, onChange]);

  const handleRemoveProductionItem = useCallback((id) => {
    const next = productionItems.filter(item => item.id !== id);
    onChange('pricing_production_items', normalizeProductionItems(next));
  }, [productionItems, onChange]);

  const productionNetValue = totals?.productionNet ?? Math.max(productionTotalValue - productionDiscountValueNumber, 0);
  const productionDiscountValueDisplay = totals?.productionDiscountValue ?? productionDiscountValueNumber;
  const singerDiscountValueNumber = totals?.singerDiscountValue ?? (Number(formState.pricing_discount_value) || 0);
  const customFeesNumber = totals?.custom ?? (Number(formState.pricing_custom_fees) || 0);
  const singerNetValue = totals?.singerNet ?? (Number(formState.ahmen_fee) || 0);
  const totalValue = totals?.total ?? (singerNetValue + productionNetValue);

  return (
    <div className="space-y-6">
      <section className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <span className="text-sm font-medium text-slate-600">Production & external services</span>
            <p className="text-xs text-slate-500">Track external suppliers, apply markup, and include totals automatically.</p>
          </div>
        </div>
        <button
          type="button"
          onClick={handleAddProductionItem}
          className="inline-flex items-center gap-1 rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
        >
          + Add item
        </button>
        {productionItems.length ? (
          <div className="space-y-3">
            {productionItems.map(item => {
              const lineTotal = calculateProductionItemTotal(item);
              return (
                <div key={item.id} className="rounded border border-slate-200 bg-white p-3 space-y-3">
                  <div className="grid gap-2 sm:grid-cols-5">
                    <label className="sm:col-span-2 text-xs font-medium uppercase tracking-wide text-slate-500">
                      Supplier / Company
                      <input
                        type="text"
                        className="mt-1 w-full rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        value={item.name}
                        onChange={event => handleProductionChange(item.id, 'name', event.target.value)}
                      />
                    </label>
                    <label className="sm:col-span-2 text-xs font-medium uppercase tracking-wide text-slate-500">
                      Description
                      <input
                        type="text"
                        className="mt-1 w-full rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        value={item.description}
                        onChange={event => handleProductionChange(item.id, 'description', event.target.value)}
                      />
                    </label>
                    <label className="text-xs font-medium uppercase tracking-wide text-slate-500">
                      Cost (£)
                      <input
                        type="number"
                        step="0.01"
                        className="mt-1 w-full rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        value={item.cost}
                        onChange={event => handleProductionChange(item.id, 'cost', event.target.value)}
                      />
                    </label>
                    <label className="text-xs font-medium uppercase tracking-wide text-slate-500">
                      Markup (%)
                      <input
                        type="number"
                        step="0.1"
                        className="mt-1 w-full rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        value={item.markup}
                        onChange={event => handleProductionChange(item.id, 'markup', event.target.value)}
                      />
                    </label>
                    <div className="flex flex-col justify-between">
                      <div>
                        <div className="text-xs font-medium uppercase tracking-wide text-slate-500">Line total</div>
                        <div className="text-sm font-semibold text-slate-700">{toCurrency(lineTotal)}</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleRemoveProductionItem(item.id)}
                        className="self-end text-xs font-medium text-red-600 hover:text-red-500"
                      >
                        Remove
                      </button>
                    </div>
                  </div>
                  <label className="text-xs font-medium uppercase tracking-wide text-slate-500">
                    Notes
                    <textarea
                      rows={2}
                      className="mt-1 w-full rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      value={item.notes}
                      onChange={event => handleProductionChange(item.id, 'notes', event.target.value)}
                    />
                  </label>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="rounded border border-dashed border-slate-300 bg-slate-50 px-3 py-2 text-sm text-slate-500">
            No production items yet. Add suppliers or services to include third-party costs.
          </div>
        )}
      </section>

      <div className="rounded border border-slate-200 bg-white p-3 text-sm space-y-2">
        <div className="flex items-center justify-between">
          <span className="font-medium text-slate-600">Production discount</span>
          {productionDiscountType === 'percent' && productionDiscountValueDisplay > 0 ? (
            <span className="text-xs text-slate-500">≈ {toCurrency(productionDiscountValueDisplay)}</span>
          ) : null}
        </div>
        <div className="flex gap-1 w-32 sm:w-36">
          {['amount', 'percent'].map(type => (
            <button
              key={type}
              type="button"
              onClick={() => {
                if (type !== productionDiscountType) onChange('pricing_production_discount_type', type);
              }}
              className={`inline-flex flex-1 items-center justify-center rounded-full border px-2.5 py-1 text-xs font-medium transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${
                type === productionDiscountType ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:border-indigo-200 hover:text-indigo-600'
              }`}
            >
              {type === 'amount' ? 'Amount (£)' : 'Percent (%)'}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          <div className="relative w-32 sm:w-36">
            <span className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-xs text-slate-400">
              {productionDiscountType === 'amount' ? '£' : '%'}
            </span>
            <input
              type="number"
              step="0.01"
              className="w-full rounded border border-slate-300 px-6 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              value={formState.pricing_production_discount || ''}
              onChange={event => onChange('pricing_production_discount', event.target.value)}
            />
          </div>
          {productionDiscountType === 'percent' ? (
            <span className="text-xs text-slate-500">≈ {toCurrency(productionDiscountValueDisplay)}</span>
          ) : null}
        </div>
      </div>

      <div className="flex items-center justify-between text-sm text-slate-600">
        <span>Total production</span>
        <span className="font-semibold text-slate-700">{toCurrency(productionNetValue)}</span>
      </div>

      <div className="rounded-lg bg-indigo-50 p-3 text-sm text-indigo-700">
        <div className="font-semibold">Quote summary</div>
        <div>{totals?.singerCount ?? 0} singer{(totals?.singerCount ?? 0) === 1 ? '' : 's'} selected · Base fee {toCurrency(totals?.base ?? 0)}</div>
        <div>Singer fees after discount: {toCurrency(totals?.singerNet ?? singerNetValue)}</div>
        <div>Production after discount: {toCurrency(totals?.productionNet ?? productionNetValue)}</div>
        <div>Singer discount: -{toCurrency(singerDiscountValueNumber)}</div>
        <div>Production discount: -{toCurrency(productionDiscountValueDisplay)}</div>
        <div>Custom fees: {toCurrency(customFeesNumber)}</div>
        <div className="font-semibold text-indigo-900">Total after adjustments: {toCurrency(totalValue)}</div>
      </div>
    </div>
  );
}

function Field({ label, name, type = 'text', value, onChange, readOnly, hint, rows = 3, step, component, options, autoFocus }) {
  const common = {
    className: 'mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500',
    value: value ?? '',
    onChange: (event) => onChange(event.target.value),
    readOnly,
    disabled: readOnly,
    step,
    name,
    autoFocus
  };

  let input;
  if (component === 'statusSelect') {
    const current = normalizeStatus(value) || 'enquiry';
    const list = options || STATUS_OPTIONS;
    input = (
      <div className="mt-1 flex flex-wrap gap-2">
        {list.map(opt => {
          const active = opt.value === current;
          const base = 'inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-sm font-medium transition focus:outline-none focus:ring-2 focus:ring-indigo-500';
          const activeStyles = STATUS_STYLES[opt.value] || 'bg-slate-200 text-slate-700 border-slate-300';
          const inactiveStyles = 'bg-white border-slate-200 text-slate-600 hover:border-indigo-200 hover:text-indigo-600';
          return (
            <button
              key={opt.value}
              type="button"
              className={`${base} ${active ? activeStyles : inactiveStyles}`}
              onClick={() => {
                if (!active) onChange(opt.value);
              }}
            >
              {opt.label}
            </button>
          );
        })}
      </div>
    );
  } else if (component === 'venueSearch') {
    const venueQueryParts = [];
    if (options?.useClient) {
      // alternate mode if we ever want client search
    }
    input = (
      <div className="mt-1 flex flex-wrap gap-2">
        <button
          type="button"
          className="inline-flex items-center gap-1.5 rounded border px-2.5 py-1 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:border-indigo-200 focus:outline-none focus:ring-2 focus:ring-indigo-500"
          onClick={() => {
            const name = document.querySelector('input[name="venue_name"]')?.value || '';
            const town = document.querySelector('input[name="venue_town"]')?.value || '';
            const postcode = document.querySelector('input[name="venue_postcode"]')?.value || '';
            const query = [name, town, postcode].filter(Boolean).join(' ');
            const url = `https://www.google.com/search?q=${encodeURIComponent(query || 'venue')}`;
            window.api?.openExternal?.(url) || window.open(url, '_blank');
          }}
        >
          Search Google
        </button>
        <button
          type="button"
          className="inline-flex items-center gap-1.5 rounded border px-2.5 py-1 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:border-indigo-200 focus:outline-none focus:ring-2 focus:ring-indigo-500"
          onClick={() => {
            const name = document.querySelector('input[name="venue_name"]')?.value || '';
            const town = document.querySelector('input[name="venue_town"]')?.value || '';
            const postcode = document.querySelector('input[name="venue_postcode"]')?.value || '';
            const query = [name, town, postcode].filter(Boolean).join(' ');
            const url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query || 'venue')}`;
            window.api?.openExternal?.(url) || window.open(url, '_blank');
          }}
        >
          Search Maps
        </button>
      </div>
    );
  } else if (type === 'textarea') {
    input = <textarea {...common} rows={rows} />;
  } else if (type === 'checkbox') {
    input = (
      <input
        type="checkbox"
        className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
        checked={Boolean(value)}
        onChange={(event) => onChange(event.target.checked)}
      />
    );
  } else if (type === 'time') {
    // Text-based 24-hour input with normalization on blur to HH:MM
    const to24h = (h, min, ap) => {
      let hour = Number(h);
      let m = Number(min);
      if (Number.isNaN(hour)) hour = 0;
      if (Number.isNaN(m)) m = 0;
      hour = Math.max(0, Math.min(23, hour));
      m = Math.max(0, Math.min(59, m));
      if (ap) {
        const ampm = ap.toUpperCase();
        hour = hour % 12;
        if (ampm === 'PM') hour += 12;
      }
      return `${String(hour).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    };
    const normalizeTime24 = (input) => {
      const raw = (input || '').toString().trim();
      if (!raw) return '';
      let s = raw.replace(/\./g, ':').replace(/-/g, ':').replace(/\s+/g, ' ').trim();
      // 12h with optional minutes and optional space before am/pm
      let m = s.match(/^(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])$/);
      if (m) return to24h(m[1], m[2] ?? '0', m[3]);
      // 24h HH:MM
      m = s.match(/^(\d{1,2}):(\d{1,2})$/);
      if (m) return to24h(m[1], m[2]);
      // Compact 3-4 digits e.g. 730 or 1530
      m = s.match(/^(\d{3,4})$/);
      if (m) {
        const num = m[1];
        const mm = num.slice(-2);
        const hh = num.slice(0, num.length - 2);
        return to24h(hh, mm);
      }
      // Bare hour
      m = s.match(/^(\d{1,2})$/);
      if (m) return to24h(m[1], '0');
      return raw; // leave unrecognized as-is
    };
    input = (
      <input
        type="text"
        placeholder="e.g. 19:30"
        {...common}
        onChange={(e) => onChange(e.target.value)}
        onBlur={(e) => onChange(normalizeTime24(e.target.value))}
      />
    );
  } else {
    input = <input type={type} {...common} />;
  }

  return (
    <label className="block text-sm font-medium text-slate-600">
      <span className="flex items-center gap-2">
        {label}
        {readOnly ? <span className="text-xs font-normal text-slate-400">auto</span> : null}
      </span>
      <div className={type === 'checkbox' ? 'mt-2' : ''}>
        {input}
      </div>
      {hint ? <p className="mt-1 text-xs text-slate-400">{hint}</p> : null}
    </label>
  );
}

function SavedVenueSelector({
  label,
  value,
  venues,
  onSelect,
  onSaveCurrent,
  onCreateNew,
  onEdit,
  onDelete,
  saving
}) {
  return (
    <div className="space-y-2">
      <label className="block text-sm font-medium text-slate-600">
        <span className="flex items-center gap-2">{label}</span>
        <select
          className='mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500'
          value={value || ''}
          onChange={event => onSelect(event.target.value)}
        >
          <option value="">Select saved venue…</option>
          {venues.map(venue => (
            <option key={venue.venue_id} value={venue.venue_id}>
              {venue.name || 'Untitled venue'}
            </option>
          ))}
        </select>
      </label>
      <div className="flex flex-wrap gap-2">
        <button
          type="button"
          onClick={onSaveCurrent}
          disabled={saving}
          className="inline-flex items-center rounded bg-slate-800 px-3 py-2 text-xs font-medium text-white hover:bg-slate-700 disabled:opacity-60 disabled:cursor-not-allowed"
        >
          {saving ? 'Saving…' : 'Save current venue'}
        </button>
        <button
          type="button"
          onClick={onEdit}
          disabled={saving || !value}
          className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60 disabled:cursor-not-allowed"
        >
          Edit selected
        </button>
        <button
          type="button"
          onClick={onDelete}
          disabled={saving || !value}
          className="inline-flex items-center rounded border border-red-200 px-3 py-2 text-xs font-medium text-red-600 hover:bg-red-50 disabled:opacity-60 disabled:cursor-not-allowed"
        >
          Delete selected
        </button>
        <button
          type="button"
          onClick={onCreateNew}
          className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
        >
          + New venue
        </button>
      </div>
    </div>
  );
}

function JobsheetEditor({
  business,
  businessId,
  jobsheetId,
  formState,
  onChange,
  onDelete,
  saving,
  deleting,
  hasExisting,
  venues,
  setVenues,
  onSaveVenue,
  venueSaving,
  setVenueSaving,
  pricingConfig,
  pricingTotals,
  onUpdateSingerPool,
  activeGroupKey: activeGroupKeyProp,
  onActiveGroupChange,
  groups,
  documents,
  documentsLoading,
  documentsError,
  documentDefinitions,
  definitionsLoading,
  onRefreshDocuments,
  onGenerateDocument,
  onExportPdf,
  onOpenDocumentFile,
  onRevealDocument,
  onDeleteDocument,
  documentFolder
}) {
  const handleFieldChange = (name, value) => {
    onChange(prev => {
      const next = applyDerivedFields({ ...prev, [name]: value });
      return next;
    });
  };

  const resolvedGroups = useMemo(() => (
    Array.isArray(groups) && groups.length ? groups : FALLBACK_JOBSHEET_GROUPS
  ), [groups]);

  // Keep refs to each rendered section to support scroll-to-section behavior
  const sectionRefs = useRef({});
  // Dynamic offset for sidebar sticky position and scroll alignment
  const [stickyTop, setStickyTop] = useState(120);
  // Per-section collapse state
  const collapsedStorageKey = useMemo(() => (
    `jobsheetEditor:collapsed:${Number(businessId) || 0}:${Number(jobsheetId) || 0}`
  ), [businessId, jobsheetId]);
  const [collapsedMap, setCollapsedMap] = useState(() => {
    try {
      const raw = typeof window !== 'undefined'
        ? (window.sessionStorage.getItem(collapsedStorageKey) || window.localStorage.getItem(collapsedStorageKey) || '')
        : '';
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_) {
      return {};
    }
  });
  const isGroupCollapsed = useCallback((key) => Boolean(collapsedMap?.[key]), [collapsedMap]);
  const toggleGroup = useCallback((key) => {
    setCollapsedMap(prev => ({ ...prev, [key]: !prev?.[key] }));
  }, []);
  const ensureExpanded = useCallback((key) => {
    setCollapsedMap(prev => (prev?.[key] ? { ...prev, [key]: false } : prev));
  }, []);
  useEffect(() => {
    try {
      if (typeof window !== 'undefined') {
        window.sessionStorage.setItem(collapsedStorageKey, JSON.stringify(collapsedMap || {}));
        if (window.localStorage.getItem('app:persistUiState') === 'true') {
          window.localStorage.setItem(collapsedStorageKey, JSON.stringify(collapsedMap || {}));
        }
      }
    } catch (_) {}
  }, [collapsedStorageKey, collapsedMap]);
  // Keep collapse map aligned with available groups
  useEffect(() => {
    setCollapsedMap(prev => {
      const allowed = new Set(resolvedGroups.map(g => g.key));
      const next = {};
      Object.keys(prev || {}).forEach(k => { if (allowed.has(k)) next[k] = prev[k]; });
      return next;
    });
  }, [resolvedGroups]);
  useEffect(() => {
    const measure = () => {
      try {
        const el = document.getElementById('jobsheet-sticky-header');
        const h = el ? (el.getBoundingClientRect().height || 0) : 0;
        setStickyTop(Math.max(0, Math.round(h + 16)));
      } catch (_) {
        setStickyTop(120);
      }
    };
    measure();
    window.addEventListener('resize', measure);
    window.addEventListener('orientationchange', measure);
    return () => {
      window.removeEventListener('resize', measure);
      window.removeEventListener('orientationchange', measure);
    };
  }, []);

  // Highlight sidebar entry based on scroll position
  const scrollRafRef = useRef(null);
  const changeByScrollRef = useRef(false);
  const detectActiveGroup = useCallback(() => {
    if (!resolvedGroups.length) return null;
    let bestKey = null;
    let bestAbs = Infinity;
    const offset = stickyTop + 8; // ensure some breathing space under sticky header
    for (const group of resolvedGroups) {
      const el = sectionRefs.current?.[group.key];
      if (!el) continue;
      const y = (el.getBoundingClientRect().top || 0) - offset;
      const score = Math.abs(y);
      if (score < bestAbs) {
        bestAbs = score;
        bestKey = group.key;
      }
    }
    return bestKey;
  }, [resolvedGroups, stickyTop]);

  useEffect(() => {
    const onScroll = () => {
      if (scrollRafRef.current != null) return;
      scrollRafRef.current = window.requestAnimationFrame(() => {
        scrollRafRef.current = null;
        const nextKey = detectActiveGroup();
        if (nextKey && nextKey !== activeGroupKeyProp) {
          changeByScrollRef.current = true;
          onActiveGroupChange?.(nextKey);
        }
      });
    };
    window.addEventListener('scroll', onScroll, { passive: true });
    return () => {
      window.removeEventListener('scroll', onScroll);
      if (scrollRafRef.current != null) {
        window.cancelAnimationFrame(scrollRafRef.current);
        scrollRafRef.current = null;
      }
    };
  }, [detectActiveGroup, activeGroupKeyProp, onActiveGroupChange]);

  const [savedVenueId, setSavedVenueId] = useState(() => (
    formState.venue_id ? String(formState.venue_id) : ''
  ));

  useEffect(() => {
    setSavedVenueId(formState.venue_id ? String(formState.venue_id) : '');
  }, [formState.venue_id]);

  const [showVenueModal, setShowVenueModal] = useState(false);
  const [venueDraft, setVenueDraft] = useState(() => buildVenueDraft());
  const [venueSearchUrl, setVenueSearchUrl] = useState('');
  const [addrQuery, setAddrQuery] = useState('');
  const [addrResults, setAddrResults] = useState([]);
  const [addrLoading, setAddrLoading] = useState(false);
  const [addrError, setAddrError] = useState('');
  const [addrPaste, setAddrPaste] = useState('');
  const addrTimerRef = useRef(null);
  const addrLastFetchRef = useRef(0);

  const parsePastedAddress = useCallback((text) => {
    try {
      const raw = (text || '').toString().trim();
      if (!raw) return null;
      const lines = raw.split(/\n|,/).map(s => s.trim()).filter(Boolean);
      // UK postcode pattern (also works broadly for UK-like codes)
      const postcodeMatch = raw.match(/([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})/i);
      const postcode = postcodeMatch ? postcodeMatch[1].toUpperCase().replace(/\s+/, ' ') : '';
      // Town: last token without postcode, or second-to-last line
      let town = '';
      if (postcode) {
        const postIndex = lines.findIndex(l => l.toUpperCase().includes(postcode));
        if (postIndex > 0) {
          town = lines[postIndex - 1];
        }
      }
      if (!town && lines.length >= 2) {
        town = lines[lines.length - 2];
      }
      const name = lines[0] || '';
      const address1 = lines.length > 1 ? lines[1] : '';
      const address2 = lines.length > 2 ? lines[2] : '';
      const address3 = lines.length > 3 ? lines[3] : '';
      return { name, address1, address2, address3, town, postcode };
    } catch (_) {
      return null;
    }
  }, []);

  const mapNominatimToVenue = useCallback((res) => {
    if (!res) return null;
    const addr = res.address || {};
    const name = res.name || addr.amenity || addr.building || addr.place || '';
    const address1 = [addr.house_number, addr.road].filter(Boolean).join(' ');
    const address2 = addr.suburb || addr.neighbourhood || addr.village || addr.district || '';
    const address3 = addr.county || addr.state_district || addr.state || '';
    const town = addr.city || addr.town || addr.village || addr.hamlet || addr.municipality || addr.suburb || '';
    const postcode = addr.postcode || '';
    return { name, address1, address2, address3, town, postcode };
  }, []);

  // Debounced/throttled address search via Nominatim (respecting usage policies: low volume, debounced)
  useEffect(() => {
    if (!showVenueModal) return; // only when modal visible
    const q = (addrQuery || '').trim();
    if (addrTimerRef.current) { clearTimeout(addrTimerRef.current); addrTimerRef.current = null; }
    if (q.length < 3) { setAddrResults([]); setAddrError(''); setAddrLoading(false); return; }
    addrTimerRef.current = setTimeout(async () => {
      try {
        const now = Date.now();
        const since = now - (addrLastFetchRef.current || 0);
        const wait = since < 1100 ? (1100 - since) : 0; // ~1 req/sec
        setAddrLoading(true);
        setAddrError('');
        await new Promise(r => setTimeout(r, wait));
        const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=8&addressdetails=1&q=${encodeURIComponent(q)}`;
        const res = await fetch(url, { headers: { 'Accept-Language': 'en' } });
        addrLastFetchRef.current = Date.now();
        if (!res.ok) throw new Error(`Search failed (${res.status})`);
        const data = await res.json();
        setAddrResults(Array.isArray(data) ? data : []);
      } catch (err) {
        setAddrError(err?.message || 'Search failed');
      } finally {
        setAddrLoading(false);
      }
    }, 450);
    return () => { if (addrTimerRef.current) { clearTimeout(addrTimerRef.current); addrTimerRef.current = null; } };
  }, [addrQuery, showVenueModal]);

  // Local override for definition lock state so the inline UI updates immediately after toggle
  const [definitionLocks, setDefinitionLocks] = useState({});

  const openVenueModal = (venue = null) => {
    if (venue) {
      setVenueDraft(buildVenueDraft(venue));
    } else {
      setVenueDraft(buildVenueDraft({
        venue_id: formState.venue_id,
        name: formState.venue_name,
        address1: formState.venue_address1,
        address2: formState.venue_address2,
        address3: formState.venue_address3,
        town: formState.venue_town,
        postcode: formState.venue_postcode,
        is_private: formState.venue_same_as_client
      }));
    }
    setShowVenueModal(true);
    setVenueSearchUrl('');
  };

  const closeVenueModal = () => {
    setShowVenueModal(false);
  };

  const handleVenueDraftChange = (field, value) => {
    setVenueDraft(prev => ({ ...prev, [field]: value }));
  };

  const handleCreateVenue = async () => {
    if (venueSaving) return;
    if (!venueDraft.name.trim()) return;
    const savedId = await onSaveVenue({ ...venueDraft });
    if (!savedId) return;
    setVenues(prev => {
      const draft = buildVenueDraft({ ...venueDraft, venue_id: savedId });
      const others = prev.filter(item => Number(item.venue_id) !== Number(savedId));
      const next = [...others, draft];
      next.sort((a, b) => a.name.localeCompare(b.name));
      return next;
    });
    setSavedVenueId(String(savedId));
    setShowVenueModal(false);
  };

  const handleToggleDefinitionLockInline = useCallback(async (item) => {
    if (!item) return;
    const api = window.api;
    try {
      // PDF document lock toggle (has document_id)
      if (item.document_id != null) {
        await api?.setDocumentLock?.(item.document_id, !(item.is_locked ? 1 : 0));
        await onRefreshDocuments?.();
        return;
      }

      // Definition lock toggle
      if (!item.key) return;
      if (!api || typeof api.saveDocumentDefinition !== 'function') return;
      const nextLocked = item.is_locked ? 0 : 1;
      await api.saveDocumentDefinition(businessId, {
        key: item.key,
        doc_type: item.doc_type,
        label: item.label,
        description: item.description,
        invoice_variant: item.invoice_variant,
        template_path: item.template_path,
        is_primary: item.is_primary ? 1 : 0,
        is_active: item.is_active === 0 ? 0 : 1,
        is_locked: nextLocked,
        sort_order: item.sort_order
      });
      setDefinitionLocks(prev => ({ ...prev, [item.key]: Boolean(nextLocked) }));
    } catch (err) {
      console.warn('Inline lock toggle failed', err);
    }
  }, [businessId, onRefreshDocuments]);

  // Scroll to a group section and propagate selection upstream
  const scrollToGroup = useCallback((key) => {
    const el = sectionRefs.current?.[key];
    if (!el) return;
    const sticky = document.getElementById('jobsheet-sticky-header');
    const stickyHeight = sticky ? (sticky.getBoundingClientRect().height || 0) : stickyTop;
    const extraGap = 12;
    const top = el.getBoundingClientRect().top + window.scrollY - (stickyHeight + extraGap);
    try {
      window.scrollTo({ top: Math.max(top, 0), behavior: 'smooth' });
    } catch (_) {
      window.scrollTo(0, Math.max(top, 0));
    }
  }, [stickyTop]);

  const setGroupKey = useCallback((nextKey) => {
    if (!nextKey) return;
    if (!resolvedGroups.some(group => group.key === nextKey)) return;
    onActiveGroupChange?.(nextKey);
    ensureExpanded(nextKey);
    // Defer to ensure refs exist
    setTimeout(() => scrollToGroup(nextKey), 0);
  }, [resolvedGroups, onActiveGroupChange, scrollToGroup, ensureExpanded]);

  // When an external action requests a specific section (e.g., 'documents'),
  // honor it by scrolling into view.
  const initialSectionAppliedRef = useRef(false);
  const lastProgrammaticKeyRef = useRef(null);
  useEffect(() => {
    if (!activeGroupKeyProp) return;
    const key = String(activeGroupKeyProp);
    if (!resolvedGroups.some(g => g.key === key)) return;
    ensureExpanded(key);
    // Suppress the very first auto-scroll on initial mount to avoid jumping the page
    if (!initialSectionAppliedRef.current) {
      initialSectionAppliedRef.current = true;
      // Remember the first key so a subsequent identical programmatic set doesn't trigger scroll
      lastProgrammaticKeyRef.current = key;
      return;
    }
    // If the requested key hasn't changed, don't scroll again
    if (lastProgrammaticKeyRef.current === key) {
      return;
    }
    lastProgrammaticKeyRef.current = key;
    if (changeByScrollRef.current) {
      // Skip programmatic scroll when the change came from natural scrolling
      changeByScrollRef.current = false;
    } else {
      // Skip if user is currently editing any input in the editor
      const ae = document.activeElement;
      if (ae && (/^(input|textarea|select)$/i).test(ae.tagName)) {
        return;
      }
      // Defer for layout
      setTimeout(() => scrollToGroup(key), 0);
    }
  }, [activeGroupKeyProp, resolvedGroups, scrollToGroup, ensureExpanded]);

  const handleSelectSavedVenue = (venueIdValue) => {
    const value = venueIdValue || '';
    setSavedVenueId(value);
    if (!value) {
      handleFieldChange('venue_id', null);
      return;
    }
    const venue = venues.find(v => String(v.venue_id) === value);
    if (!venue) return;
    handleFieldChange('venue_same_as_client', false);
    handleFieldChange('venue_id', venue.venue_id);
    handleFieldChange('venue_name', venue.name || '');
    handleFieldChange('venue_address1', venue.address1 || '');
    handleFieldChange('venue_address2', venue.address2 || '');
    handleFieldChange('venue_address3', venue.address3 || '');
    handleFieldChange('venue_town', venue.town || '');
    handleFieldChange('venue_postcode', venue.postcode || '');
    handleFieldChange('venue_same_as_client', Boolean(venue.is_private));
  };

  const handleEditSavedVenue = () => {
    if (!savedVenueId) return;
    const venue = venues.find(v => String(v.venue_id) === savedVenueId);
    if (!venue) return;
    openVenueModal(venue);
  };

  // Keep status in Client Details (editor) and Jobsheet list in sync
  useEffect(() => {
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload || payload.businessId !== businessId) return;
      if (payload.type !== 'jobsheet-updated') return;
      const payloadId = payload.jobsheetId != null
        ? Number(payload.jobsheetId)
        : (payload.snapshot && payload.snapshot.jobsheet_id != null ? Number(payload.snapshot.jobsheet_id) : null);
      if (payloadId == null || (jobsheetId != null && Number(jobsheetId) !== payloadId)) return;
      const nextStatus = normalizeStatus(payload.snapshot?.status) || null;
      if (!nextStatus) return;
      onChange(prev => {
        const current = normalizeStatus(prev?.status) || '';
        if (current === nextStatus) return prev; // no change; avoid save loop
        return { ...prev, status: nextStatus };
      });
    });
    return () => unsubscribe?.();
  }, [businessId, jobsheetId, onChange]);

  const handleDeleteSavedVenue = async () => {
    if (!savedVenueId) return;
    const venue = venues.find(v => String(v.venue_id) === savedVenueId);
    if (!venue) return;
    const confirmed = window.confirm(`Delete venue "${venue.name || 'Untitled venue'}"? This cannot be undone.`);
    if (!confirmed) return;
    const api = window.api;
    if (!api || !api.deleteAhmenVenue) {
      setError('Unable to delete venue: API unavailable');
      return;
    }
    setVenueSaving(true);
    try {
      await api.deleteAhmenVenue(Number(venue.venue_id));
      setVenues(prev => prev.filter(item => Number(item.venue_id) !== Number(venue.venue_id)));
      setSavedVenueId('');
      setFormState(prev => {
        if (prev.venue_id !== venue.venue_id) return prev;
        return applyDerivedFields({
          ...prev,
          venue_id: null,
          venue_name: '',
          venue_address1: '',
          venue_address2: '',
          venue_address3: '',
          venue_town: '',
          venue_postcode: '',
          venue_same_as_client: false
        });
      });
      const updatedVenues = await api.getAhmenVenues({ businessId });
      setVenues(normalizeVenues(updatedVenues));
      setMessage('Venue deleted');
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to delete venue', err);
      setError(err?.message || 'Unable to delete venue');
    } finally {
      setVenueSaving(false);
    }
  };

  

  return (
    <>
      <div className="space-y-6">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
          <div>
            <h2 className="text-xl font-semibold text-slate-800">{hasExisting ? 'Edit jobsheet' : 'New jobsheet'}</h2>
            <p className="text-sm text-slate-500">Business: {business.business_name}</p>
          </div>
          <div className="flex flex-col items-stretch gap-3 sm:flex-row sm:items-center sm:gap-3">
            {hasExisting ? (
              <button
                onClick={onDelete}
                disabled={deleting}
                className="text-sm font-medium text-red-600 hover:text-red-500 disabled:opacity-60"
              >
                {deleting ? 'Deleting…' : 'Delete jobsheet'}
              </button>
            ) : null}
          </div>
        </div>

      <div className="flex flex-col gap-6 lg:flex-row">
        <nav className="lg:w-64 flex-shrink-0 lg:sticky self-start" style={{ top: stickyTop }}>
          <div className="space-y-2" role="navigation" aria-orientation="vertical" aria-label="Jump to section">
            {resolvedGroups.map(group => {
              const isActive = activeGroupKeyProp && group.key === activeGroupKeyProp;
              const icon = group.icon ?? getGroupIcon(group.key);
              return (
                <button
                  key={group.key}
                  type="button"
                  onClick={() => setGroupKey(group.key)}
                  aria-current={isActive ? 'page' : undefined}
                  className={`group flex w-full items-center gap-3 rounded-lg border px-3 py-3 text-left transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${isActive ? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold shadow-sm' : 'border-transparent bg-white text-slate-600 hover:bg-slate-50 hover:border-slate-200'}`}
                >
                  <span className={`flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-full text-lg transition ${isActive ? 'bg-indigo-100 text-indigo-700 shadow-sm' : 'bg-slate-100 text-slate-500 group-hover:bg-slate-200 group-hover:text-slate-700'}`}>
                    {icon}
                  </span>
                  <span className="flex-1">
                    <span className="block text-sm font-semibold">{group.title}</span>
                    {group.description ? (
                      <span className="mt-1 block text-xs text-slate-500">{group.description}</span>
                    ) : null}
                  </span>
                </button>
              );
            })}
          </div>
        </nav>

        <div className="flex-1">
          {resolvedGroups.length ? (
            resolvedGroups.map(group => (
              <section
                key={group.key}
                id={`jobsheet-section-${group.key}`}
                ref={el => { if (el) sectionRefs.current[group.key] = el; }}
                className="bg-white border border-slate-200 rounded-lg p-5 space-y-5"
                style={{ scrollMarginTop: stickyTop + 8 }}
              >
                <div>
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <h3 className="text-lg font-semibold text-slate-700">{group.title}</h3>
                      {group.key === 'documents' ? (
                        <>
                          <button
                            type="button"
                            onClick={async () => {
                              try {
                                const res = await window.api?.ensureJobsheetFolder?.({ businessId, jobsheetId, jobsheetSnapshot: formState });
                                const folderPath = res?.folder_path || res?.path || '';
                                if (!folderPath) throw new Error('Unable to resolve folder path');
                                const open = await window.api?.openPath?.(folderPath);
                                if (open && open.ok === false) throw new Error(open.message || 'Unable to open folder');
                                await onRefreshDocuments?.();
                              } catch (err) {
                                window.alert(err?.message || 'Unable to open job folder');
                              }
                            }}
                            className="inline-flex items-center gap-1.5 rounded bg-indigo-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          >
                            <span aria-hidden>📂</span>
                            <span>Open job folder</span>
                          </button>
                        </>
                      ) : null}
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        aria-expanded={!isGroupCollapsed(group.key)}
                        aria-controls={`jobsheet-section-${group.key}`}
                        onClick={() => toggleGroup(group.key)}
                        className="inline-flex items-center gap-1.5 rounded border border-slate-200 px-2.5 py-1 text-xs font-medium text-slate-600 hover:text-indigo-600 hover:border-indigo-200 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        title={isGroupCollapsed(group.key) ? 'Expand section' : 'Collapse section'}
                      >
                        <span aria-hidden>{isGroupCollapsed(group.key) ? '▸' : '▾'}</span>
                        <span className="hidden sm:inline">{isGroupCollapsed(group.key) ? 'Expand' : 'Collapse'}</span>
                      </button>
                    </div>
                  </div>
                  {group.description && !isGroupCollapsed(group.key) ? (
                    <p className="mt-1 text-sm text-slate-500">{group.description}</p>
                  ) : null}
                </div>
                {!isGroupCollapsed(group.key) ? (
                  <div className="space-y-4">
                    {group.fields.map(field => {
                      if (field.component === 'pricingPanel') {
                        return pricingConfig ? (
                          <PricingPanel
                            key={field.name}
                          pricingConfig={pricingConfig}
                          pricingTotals={pricingTotals}
                          formState={formState}
                          onChange={handleFieldChange}
                          hasExisting={hasExisting}
                          onUpdateSingerPool={onUpdateSingerPool}
                        onFocusPricingPanel={() => setGroupKey('pricing')}
                        />
                      ) : (
                        <div key={field.name} className="rounded border border-slate-200 bg-white p-4 text-sm text-slate-500">
                          Loading pricing configuration…
                        </div>
                      );
                    }
                    if (field.component === 'gigInfoPanel') {
                      return (
                        <GigInfoPanel
                          key="gigInfoPanel"
                          formState={formState}
                          onChange={handleFieldChange}
                          businessId={Number(businessId) || 0}
                          jobsheetId={jobsheetId}
                        />
                      );
                    }
                    if (field.component === 'productionPanel') {
                      return (
                        <ProductionPanel
                          key={field.name}
                          formState={formState}
                          onChange={handleFieldChange}
                          totals={pricingTotals}
                        />
                      );
                    }
                    if (field.component === 'documentsPanel') {
                      return (
                        <DocumentsInlinePanel
                          key="documentsPanel"
                          jobsheetId={jobsheetId}
                          jobsheetStatus={formState.status}
                          documentDefinitions={documentDefinitions}
                          documents={documents}
                          loading={documentsLoading}
                          definitionsLoading={definitionsLoading}
                          error={documentsError}
                          onRefresh={onRefreshDocuments}
                          onGenerate={onGenerateDocument}
                          onExportPdf={onExportPdf}
                          onToggleLock={handleToggleDefinitionLockInline}
                          locksOverride={definitionLocks}
                          onOpenFile={onOpenDocumentFile}
                          onRevealFile={onRevealDocument}
                          onDelete={onDeleteDocument}
                          documentFolder={documentFolder}
                          businessId={businessId}
                          lastInvoiceNumber={business?.last_invoice_number}
                          jobsheetSnapshot={formState}
                        />
                      );
                    }
                    if (field.component === 'savedVenueSelector') {
                      return (
                        <SavedVenueSelector
                          key={field.name}
                          label={field.label}
                          value={savedVenueId}
                          venues={venues}
                          onSelect={handleSelectSavedVenue}
                          onSaveCurrent={() => onSaveVenue()}
                          onCreateNew={() => openVenueModal()}
                          onEdit={handleEditSavedVenue}
                          onDelete={handleDeleteSavedVenue}
                          saving={venueSaving}
                        />
                      );
                    }

                    const resolvedValue = field.name === 'status'
                      ? (formState.status || 'enquiry')
                      : field.type === 'checkbox'
                        ? Boolean(formState[field.name])
                        : formState[field.name] ?? '';

                     return (
                       <Field
                         key={field.name}
                         label={field.label}
                         name={field.name}
                         type={field.type || 'text'}
                         step={field.step}
                         rows={field.rows}
                         hint={field.hint}
                         readOnly={field.name === 'venue_name' ? Boolean(formState.venue_same_as_client) : field.readOnly}
                         component={field.component}
                         options={field.options}
                         autoFocus={(!hasExisting && field.name === 'client_name') ? true : undefined}
                         value={resolvedValue}
                         onChange={value => handleFieldChange(
                           field.name,
                           field.type === 'checkbox' ? Boolean(value) : value
                         )}
                       />
                     );
                  })}
                  </div>
                ) : null}
              </section>
            ))
          ) : (
            <div className="rounded-lg border border-slate-200 bg-white p-5 text-sm text-slate-500">No sections available.</div>
          )}
        </div>
      </div>

      <div className="flex items-center justify-end text-sm text-slate-500 min-h-[1.5rem]">
        {saving ? 'Saving changes…' : null}
      </div>
      </div>

      {showVenueModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4">
          <div className="w-full max-w-5xl max-h-[90vh] rounded-lg bg-white p-6 shadow-xl flex flex-col">
            <div className="flex items-start justify-between">
              <div>
                <h3 className="text-lg font-semibold text-slate-800">Add new venue</h3>
                <p className="text-sm text-slate-500">Capture the venue details and save them to reuse later.</p>
              </div>
              <button
                type="button"
                onClick={closeVenueModal}
                className="text-slate-400 hover:text-slate-600"
                aria-label="Close venue modal"
              >
                ✕
              </button>
            </div>
            <div className="mt-4 space-y-3">
              <div className="flex items-center justify-start">
                <button
                  type="button"
                  className="inline-flex items-center rounded bg-indigo-50 text-indigo-800 border border-indigo-200 px-3 py-1.5 text-xs font-medium hover:bg-indigo-100"
                  onClick={() => setVenueDraft(buildVenueDraft())}
                >
                  Clear fields
                </button>
              </div>
              <label className="block text-sm font-medium text-slate-600">
                Venue name
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={venueDraft.name}
                  onChange={event => handleVenueDraftChange('name', event.target.value)}
                />
              </label>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  className="inline-flex items-center gap-1.5 rounded border px-2.5 py-1 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:border-indigo-200 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  onClick={() => {
                    const queryParts = [venueDraft.name, venueDraft.town, venueDraft.postcode, venueDraft.address1]
                      .filter(Boolean)
                      .join(' ');
                    const fallbackParts = [formState.venue_name, formState.venue_town, formState.venue_postcode, formState.venue_address1]
                      .filter(Boolean)
                      .join(' ');
                    const q = queryParts || fallbackParts || 'venue address';
                    const url = `https://www.google.com/search?q=${encodeURIComponent(q)}`;
                    setVenueSearchUrl(url);
                  }}
                >
                  Search Google
                </button>
                <button
                  type="button"
                  className="inline-flex items-center gap-1.5 rounded border px-2.5 py-1 text-sm font-medium text-slate-600 hover:text-indigo-600 hover:border-indigo-200 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  onClick={() => {
                    const queryParts = [venueDraft.name, venueDraft.town, venueDraft.postcode, venueDraft.address1]
                      .filter(Boolean)
                      .join(' ');
                    const fallbackParts = [formState.venue_name, formState.venue_town, formState.venue_postcode, formState.venue_address1]
                      .filter(Boolean)
                      .join(' ');
                    const q = queryParts || fallbackParts || 'venue address';
                    const url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(q)}`;
                    setVenueSearchUrl(url);
                  }}
                >
                  Search Maps
                </button>
              </div>

              {/* Quick address finder (OpenStreetMap Nominatim) */}
              <div className="mt-3 rounded border border-slate-200 p-3">
                <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Quick address finder</div>
                <div className="flex flex-col gap-2 md:flex-row md:items-center">
                  <input
                    type="search"
                    placeholder="Type part of an address…"
                    className="w-full md:max-w-md rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    value={addrQuery}
                    onChange={e => setAddrQuery(e.target.value)}
                  />
                  <div className="text-xs text-slate-500 md:ml-2">Powered by OpenStreetMap</div>
                </div>
                {addrError ? (
                  <div className="mt-2 rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-700">{addrError}</div>
                ) : null}
                <div className="mt-2 max-h-64 overflow-auto">
                  {addrLoading ? (
                    <div className="px-2 py-1 text-xs text-slate-500">Searching…</div>
                  ) : (addrResults || []).length ? (
                    (addrResults || []).map((res, idx) => {
                      const addr = res.address || {};
                      const line1 = [addr.house_number, addr.road].filter(Boolean).join(' ');
                      const town = addr.city || addr.town || addr.village || addr.hamlet || addr.municipality || addr.suburb || '';
                      const postcode = addr.postcode || '';
                      const title = res.display_name || line1 || res.name || 'Address';
                      return (
                        <div key={res.place_id || idx} className="flex items-start justify-between gap-3 border-b border-slate-100 px-2 py-2 last:border-b-0">
                          <div className="min-w-0">
                            <div className="truncate text-sm font-medium text-slate-800" title={title}>{title}</div>
                            <div className="text-xs text-slate-500">{line1 || res.name || '—'}{town ? ` · ${town}` : ''}{postcode ? ` · ${postcode}` : ''}</div>
                          </div>
                          <div className="flex flex-shrink-0 gap-2">
                            <button
                              type="button"
                              className="rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 hover:bg-slate-50"
                              onClick={() => {
                                const mapped = mapNominatimToVenue(res);
                                if (mapped) setVenueDraft(prev => buildVenueDraft({ ...prev, ...mapped }));
                              }}
                            >
                              Use
                            </button>
                            <button
                              type="button"
                              className="rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 hover:bg-slate-50"
                              onClick={() => {
                                const mapped = mapNominatimToVenue(res);
                                const lines = [
                                  mapped.name,
                                  mapped.address1,
                                  mapped.address2,
                                  mapped.address3,
                                  [mapped.town, mapped.postcode].filter(Boolean).join(' ')
                                ].filter(Boolean);
                                const text = lines.join('\n');
                                try { window.api?.copyTextToClipboard?.(text); } catch (_) {}
                              }}
                            >
                              Copy
                            </button>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="px-2 py-1 text-xs text-slate-400">Enter at least 3 characters to search.</div>
                  )}
                </div>
              </div>

              {/* Paste address to auto-split */}
              <div className="mt-3 rounded border border-slate-200 p-3">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-500">Paste address</div>
                <textarea
                  rows={3}
                  className="w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  placeholder="Paste a full address here…"
                  value={addrPaste}
                  onChange={e => setAddrPaste(e.target.value)}
                />
                <div className="mt-2 flex gap-2">
                  <button
                    type="button"
                    className="rounded bg-slate-800 px-3 py-1.5 text-xs font-medium text-white hover:bg-slate-700"
                    onClick={() => {
                      const mapped = parsePastedAddress(addrPaste || '');
                      if (mapped) setVenueDraft(prev => buildVenueDraft({ ...prev, ...mapped }));
                    }}
                  >
                    Fill fields
                  </button>
                  <button
                    type="button"
                    className="rounded border border-slate-300 px-3 py-1.5 text-xs text-slate-700 hover:bg-slate-50"
                    onClick={() => setAddrPaste('')}
                  >
                    Clear
                  </button>
                </div>
              </div>
              {venueSearchUrl ? (
                <div className="mt-2 overflow-hidden rounded border border-slate-200 h-[82vh] md:h-[86vh]"
                >
                  {/* Electron webview renders external content inside the modal */}
                  <webview
                    src={venueSearchUrl}
                    allowpopups="false"
                    style={{ width: '100%', height: '100%' }}
                  />
                </div>
              ) : null}
              <label className="block text-sm font-medium text-slate-600">
                Address line 1
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={venueDraft.address1}
                  onChange={event => handleVenueDraftChange('address1', event.target.value)}
                />
              </label>
              <label className="block text-sm font-medium text-slate-600">
                Address line 2
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={venueDraft.address2}
                  onChange={event => handleVenueDraftChange('address2', event.target.value)}
                />
              </label>
              <label className="block text-sm font-medium text-slate-600">
                Address line 3
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={venueDraft.address3}
                  onChange={event => handleVenueDraftChange('address3', event.target.value)}
                />
              </label>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="block text-sm font-medium text-slate-600">
                  Town / City
                  <input
                    type="text"
                    className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                    value={venueDraft.town}
                    onChange={event => handleVenueDraftChange('town', event.target.value)}
                  />
                </label>
                <label className="block text-sm font-medium text-slate-600">
                  Postcode
                  <input
                    type="text"
                    className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                    value={venueDraft.postcode}
                    onChange={event => handleVenueDraftChange('postcode', event.target.value)}
                  />
                </label>
              </div>
              
            </div>
            <div className="mt-6 flex justify-end gap-2">
              <button
                type="button"
                onClick={closeVenueModal}
                className="inline-flex items-center rounded border border-slate-300 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleCreateVenue}
                disabled={venueSaving || !venueDraft.name.trim()}
                className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500 disabled:opacity-60 disabled:cursor-not-allowed"
              >
                {venueSaving ? 'Saving…' : 'Save venue'}
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}

function BusinessWorkspace({ business, onBusinessUpdate }) {
  const [jobsheets, setJobsheets] = useState([]);
  const [listLoading, setListLoading] = useState(true);
  const [showArchived, setShowArchived] = useState(() => {
    if (typeof window === 'undefined') return false;
    try {
      const key = `ui:${business.id}:showArchived`;
      const raw = window.localStorage.getItem(key);
      return raw === '1' || raw === 'true';
    } catch (_) {
      return false;
    }
  });
  const [sortConfig, setSortConfig] = useState(() => {
    if (typeof window === 'undefined') return { key: 'event_date', direction: 'desc' };
    try {
      const persist = window.localStorage.getItem('app:persistUiState') === 'true';
      if (!persist) return { key: 'event_date', direction: 'desc' };
      const raw = window.localStorage.getItem(`ui:${business.id}:jobsheetSort`);
      if (!raw) return { key: 'event_date', direction: 'desc' };
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object' && parsed.key) {
        const dir = parsed.direction === 'asc' ? 'asc' : (parsed.direction === 'desc' ? 'desc' : 'desc');
        return { key: String(parsed.key), direction: dir };
      }
    } catch (_err) {}
    return { key: 'event_date', direction: 'desc' };
  });
  const [deletingId, setDeletingId] = useState(null);
  const [statusUpdatingId, setStatusUpdatingId] = useState(null);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [activeJobsheetId, setActiveJobsheetId] = useState(() => {
    if (typeof window === 'undefined') return null;
    try {
      const persist = window.localStorage.getItem('app:persistUiState') === 'true';
      if (!persist) return null;
      const key = `ui:${business.id}:activeJobsheetId`;
      const raw = window.localStorage.getItem(key);
      const num = raw != null && raw !== '' ? Number(raw) : null;
      return Number.isFinite(num) ? num : null;
    } catch (_err) {
      return null;
    }
  });
  const [inlineEditorVisible, setInlineEditorVisible] = useState(() => {
    if (typeof window === 'undefined') return false;
    try {
      const persist = window.localStorage.getItem('app:persistUiState') === 'true';
      if (!persist) return false;
      const key = `ui:${business.id}:inlineVisible`;
      return window.localStorage.getItem(key) === 'true';
    } catch (_err) {
      return false;
    }
  });
  const [inlineEditorTargetId, setInlineEditorTargetId] = useState(() => {
    if (typeof window === 'undefined') return null;
    try {
      const persist = window.localStorage.getItem('app:persistUiState') === 'true';
      if (!persist) return null;
      const visibleKey = `ui:${business.id}:inlineVisible`;
      const isVisible = window.localStorage.getItem(visibleKey) === 'true';
      if (!isVisible) return null;
      const key = `ui:${business.id}:activeJobsheetId`;
      const raw = window.localStorage.getItem(key);
      const num = raw != null && raw !== '' ? Number(raw) : null;
      return Number.isFinite(num) ? num : null;
    } catch (_err) {
      return null;
    }
  });

  // Settings: Set last invoice number inline modal state
  const [setLastOpen, setSetLastOpen] = useState(false);
  const [setLastDraft, setSetLastDraft] = useState('');
  const [inlineEditorSession, setInlineEditorSession] = useState(0);
  const [updatingSavePath, setUpdatingSavePath] = useState(false);
  const [workspaceSection, setWorkspaceSection] = useState(() => {
    if (typeof window === 'undefined') return 'jobsheets';
    try {
      const stored = window.localStorage.getItem(WORKSPACE_SECTION_STORAGE_KEY);
      const match = WORKSPACE_SECTIONS.find(section => section.key === stored);
      return match ? match.key : 'jobsheets';
    } catch (err) {
      console.warn('Unable to read workspace section', err);
      return 'jobsheets';
    }
  });
  const [documents, setDocuments] = useState([]);
  const [documentsLoading, setDocumentsLoading] = useState(true);
  const [documentsError, setDocumentsError] = useState('');
  const [documentTree, setDocumentTree] = useState(null);
  const [documentTreeLoading, setDocumentTreeLoading] = useState(false);
  const [documentTreeError, setDocumentTreeError] = useState('');
  const [documentTreeCollapsed, setDocumentTreeCollapsed] = useState(() => {
    if (typeof window === 'undefined') return false;
    try {
      const stored = window.localStorage.getItem(DOCUMENT_TREE_COLLAPSE_KEY);
      if (stored === 'true') return true;
      if (stored === 'false') return false;
    } catch (err) {
      console.warn('Unable to read document tree collapse preference', err);
    }
    return false;
  });
  const [emptyingTrash, setEmptyingTrash] = useState(false);
  const [documentsGroup, setDocumentsGroup] = useState('none');
  const [documentsSearch, setDocumentsSearch] = useState('');
  const [documentColumnsState, setDocumentColumnsState] = useState(() => {
    if (typeof window === 'undefined') return { ...DEFAULT_DOCUMENT_COLUMNS_STATE };
    try {
      const stored = window.localStorage.getItem(DOCUMENT_COLUMNS_STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed && typeof parsed === 'object') {
          return {
            ...DEFAULT_DOCUMENT_COLUMNS_STATE,
            ...parsed
          };
        }
      }
    } catch (err) {
      console.warn('Unable to read document columns preference', err);
    }
    return { ...DEFAULT_DOCUMENT_COLUMNS_STATE };
  });
  const [columnsMenuOpen, setColumnsMenuOpen] = useState(false);
  const columnsMenuRef = useRef(null);
  const columnsMenuContentRef = useRef(null);
  const [columnsMenuAbove, setColumnsMenuAbove] = useState(false);
  const [selectedDocuments, setSelectedDocuments] = useState(() => new Set());
  const [showDocumentsLoading, setShowDocumentsLoading] = useState(false);
  const PERSIST_UI_KEY = 'app:persistUiState';
  const PERSIST_PREFIX = `ui:${business.id}:`;
  const [persistUi, setPersistUi] = useState(() => {
    if (typeof window === 'undefined') return false;
    try {
      return window.localStorage.getItem(PERSIST_UI_KEY) === 'true';
    } catch (_err) { return false; }
  });

  const applyStoredScroll = useCallback(() => {
    // Keep available for targeted restores, but do not auto-apply on tab switch.
    if (!persistUi || typeof window === 'undefined') return;
    try {
      const y = Number(window.localStorage.getItem(`${PERSIST_PREFIX}scrollY`) || '0');
      if (!Number.isFinite(y) || y <= 0) return;
      let attempts = 24; // ~1.2s total
      const tick = () => {
        try {
          window.scrollTo(0, y);
          if (Math.abs((window.scrollY || 0) - y) < 2 || attempts-- <= 0) return;
          setTimeout(tick, 50);
        } catch (_err) {}
      };
      setTimeout(tick, 50);
    } catch (_err) {}
  }, [persistUi, PERSIST_PREFIX]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      window.localStorage.setItem(DOCUMENT_TREE_COLLAPSE_KEY, documentTreeCollapsed ? 'true' : 'false');
    } catch (err) {
      console.warn('Unable to store document tree collapse preference', err);
    }
  }, [documentTreeCollapsed]);

  // Persist key UI state if enabled
  useEffect(() => {
    if (typeof window === 'undefined') return;
    try { window.localStorage.setItem(PERSIST_UI_KEY, persistUi ? 'true' : 'false'); } catch (_err) {}
  }, [persistUi]);

  useEffect(() => {
    if (!persistUi || typeof window === 'undefined') return;
    try { window.localStorage.setItem(`${PERSIST_PREFIX}workspaceSection`, workspaceSection); } catch (_err) {}
  }, [persistUi, workspaceSection]);

  useEffect(() => {
    if (!persistUi || typeof window === 'undefined') return;
    try { window.localStorage.setItem(`${PERSIST_PREFIX}activeJobsheetId`, activeJobsheetId != null ? String(activeJobsheetId) : ''); } catch (_err) {}
    try { window.localStorage.setItem(`${PERSIST_PREFIX}inlineVisible`, inlineEditorVisible ? 'true' : 'false'); } catch (_err) {}
  }, [persistUi, activeJobsheetId, inlineEditorVisible]);

  useEffect(() => {
    if (!persistUi || typeof window === 'undefined') return;
    const onScroll = () => {
      try { window.localStorage.setItem(`${PERSIST_PREFIX}scrollY`, String(window.scrollY || 0)); } catch (_err) {}
    };
    window.addEventListener('scroll', onScroll, { passive: true });
    return () => window.removeEventListener('scroll', onScroll);
  }, [persistUi]);

  // Debounce documentsLoading indicator to prevent UI shake on very fast operations
  useEffect(() => {
    let timer = null;
    if (documentsLoading) {
      timer = setTimeout(() => setShowDocumentsLoading(true), 180);
    } else {
      setShowDocumentsLoading(false);
    }
    return () => { if (timer) clearTimeout(timer); };
  }, [documentsLoading]);

  // Restore UI state on mount (without auto-scrolling)
  useEffect(() => {
    if (!persistUi || typeof window === 'undefined') return;
    try {
      const storedSection = window.localStorage.getItem(`${PERSIST_PREFIX}workspaceSection`);
      if (storedSection && WORKSPACE_SECTIONS.some(s => s.key === storedSection)) {
        setWorkspaceSection(storedSection);
      }
      const storedJob = window.localStorage.getItem(`${PERSIST_PREFIX}activeJobsheetId`);
      const storedVisible = window.localStorage.getItem(`${PERSIST_PREFIX}inlineVisible`) === 'true';
      if (storedJob) {
        const idNum = Number(storedJob);
        if (Number.isFinite(idNum)) {
          setActiveJobsheetId(idNum);
          if (storedVisible) {
            setInlineEditorTargetId(idNum);
          } else {
            setInlineEditorTargetId(null);
          }
          setInlineEditorVisible(storedVisible);
        }
      }
    } catch (_err) {}
  }, [persistUi, applyStoredScroll]);
  // Stop auto-scroll on tab switch; scroll only on explicit actions (open/create)

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return () => {};
    if (!window.api) return () => {};
    window.api.watchDocuments?.({ businessId: business.id }).catch(err => {
      console.warn('Unable to start documents watcher', err);
    });
    const unsubscribe = window.api.onDocumentsChange?.((payload) => {
      if (!payload || payload.businessId !== business.id) return;
      refreshDocuments();
    });
    return () => {
      unsubscribe?.();
    };
  }, [business.id, refreshDocuments]);

  const normalizeJobsheet = useCallback(item => ({
    ...item,
    status: normalizeStatus(item.status) || 'enquiry'
  }), []);

  const activeJobsheetIdRef = useRef(null);
  useEffect(() => {
    activeJobsheetIdRef.current = activeJobsheetId != null ? Number(activeJobsheetId) : null;
  }, [activeJobsheetId]);

  const mergeJobsheetSnapshot = useCallback((snapshot) => {
    if (!snapshot || snapshot.jobsheet_id == null) return;
    setJobsheets(prev => {
      let found = false;
      const next = prev.map(job => {
        if (job.jobsheet_id === snapshot.jobsheet_id) {
          found = true;
          return normalizeJobsheet({ ...job, ...snapshot });
        }
        return job;
      });
      if (!found) {
        next.push(normalizeJobsheet(snapshot));
      }
      return next;
    });
  }, [normalizeJobsheet]);

  const refreshJobsheets = useCallback(async () => {
    setListLoading(true);
    try {
      const api = window.api;
      if (!api || !api.getAhmenJobsheets) {
        setError('Unable to load jobsheets: API unavailable');
        setListLoading(false);
        return;
      }
      const data = await api.getAhmenJobsheets({ businessId: business.id, includeArchived: showArchived });
      const mapped = (data || []).map(normalizeJobsheet);
      setJobsheets(mapped);

      const currentActive = activeJobsheetIdRef.current;
      if (currentActive != null) {
        const exists = mapped.some(job => job?.jobsheet_id != null && Number(job.jobsheet_id) === currentActive);
        if (exists) {
          setActiveJobsheetId(currentActive);
          // Preserve current visibility; only retarget when already visible
          setInlineEditorTargetId(prev => (inlineEditorVisible ? currentActive : prev));
          // Do not force visibility true here; keep user’s last state
        }
      }
    } catch (err) {
      console.error('Failed to refresh jobsheets', err);
      setError(err?.message || 'Unable to refresh jobsheets');
    } finally {
      setListLoading(false);
    }
  }, [business.id, normalizeJobsheet, showArchived, inlineEditorVisible]);

  const loadDocumentTree = useCallback(async () => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentTree(null);
      setDocumentTreeLoading(false);
      setDocumentTreeError('');
      return;
    }
    setDocumentTreeLoading(true);
    setDocumentTreeError('');
    try {
      const api = window.api;
      if (!api || typeof api.listDocumentTree !== 'function') {
        throw new Error('Document tree unavailable');
      }
      const tree = await api.listDocumentTree({ businessId: business.id });
      setDocumentTree(tree || null);
    } catch (err) {
      console.error('Failed to load document tree', err);
      setDocumentTreeError(err?.message || 'Unable to load document tree');
      setDocumentTree(null);
    } finally {
      setDocumentTreeLoading(false);
    }
  }, [business.id]);

  const refreshDocuments = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDocuments([]);
      setDocumentsLoading(false);
      setDocumentsError('');
      return;
    }
    setDocumentsLoading(true);
    setDocumentsError('');
    try {
      if (DOCUMENT_FEATURES_ENABLED || DOCUMENT_GENERATION_ENABLED) {
        const api = window.api;
        if (!api || typeof api.listJobsheetDocuments !== 'function') {
          throw new Error('Unable to load documents: API unavailable');
        }
        const response = await api.listJobsheetDocuments({ businessId: business.id });
        const docs = Array.isArray(response?.documents) ? response.documents : [];
        setDocuments(docs);
      }
      if (DOCUMENT_FEATURES_ENABLED) {
        await loadDocumentTree();
      } else {
        setDocumentTree(null);
      }
    } catch (err) {
      console.error('Failed to refresh documents', err);
      setDocumentsError(err?.message || 'Unable to load documents');
    } finally {
      setDocumentsLoading(false);
    }
  }, [business.id, loadDocumentTree]);

  const handleRefreshDocuments = useCallback(() => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) return;
    refreshDocuments();
  }, [refreshDocuments]);

  const handleOpenDocumentsFolder = useCallback(async () => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    setDocumentsError('');
    if (!business.save_path) {
      setDocumentsError('Documents folder not configured');
      return;
    }
    try {
      const response = await window.api?.openPath?.(business.save_path);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open documents folder');
      }
    } catch (err) {
      console.error('Failed to open documents folder', err);
      setDocumentsError(err?.message || 'Unable to open documents folder');
    }
  }, [business.save_path]);

  const handleOpenDocumentFile = useCallback(async (filePath) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    setDocumentsError('');
    if (!filePath) {
      setDocumentsError('Document file not available');
      return;
    }
    try {
      const response = await window.api?.openPath?.(filePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open document');
      }
    } catch (err) {
      console.error('Failed to open document', err);
      setDocumentsError(err?.message || 'Unable to open document');
    }
  }, []);

  const handleOpenTreeNode = useCallback(async (node) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    if (!node?.absolutePath) return;
    try {
      setDocumentsError('');
      const response = await window.api?.openPath?.(node.absolutePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open path');
      }
    } catch (err) {
      console.error('Failed to open path', err);
      setDocumentsError(err?.message || 'Unable to open path');
    }
  }, []);

  const handleRevealTreeNode = useCallback(async (node) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    if (!node?.absolutePath) return;
    try {
      setDocumentsError('');
      const response = await window.api?.showItemInFolder?.(node.absolutePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to reveal path');
      }
    } catch (err) {
      console.error('Failed to reveal path', err);
      setDocumentsError(err?.message || 'Unable to reveal path');
    }
  }, []);

  const handleDeleteTreeFolder = useCallback(async (node) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    if (!node?.path) {
      setDocumentsError('Cannot delete the root documents folder.');
      return;
    }
    const confirmed = window.confirm(`Move folder "${node.name}" to trash?`);
    if (!confirmed) return;
    try {
      setDocumentsError('');
      await window.api?.deleteDocumentFolder?.({ businessId: business.id, relativePath: node.path });
      setMessage('Folder moved to trash');
      await refreshDocuments();
      await loadDocumentTree();
      setTimeout(() => setMessage(''), 2000);
    } catch (err) {
      console.error('Failed to delete folder', err);
      setDocumentsError(err?.message || 'Unable to delete folder');
    }
  }, [business.id, refreshDocuments, loadDocumentTree]);

  const handleDeleteTreeFile = useCallback(async (node) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    if (!node?.absolutePath) return;
    const confirmed = window.confirm(`Move file "${node.name}" to trash?`);
    if (!confirmed) return;
    try {
      setDocumentsError('');
      await window.api?.deleteDocumentByPath?.({ businessId: business.id, absolutePath: node.absolutePath });
      setMessage('File moved to trash');
      await refreshDocuments();
      await loadDocumentTree();
      setTimeout(() => setMessage(''), 2000);
    } catch (err) {
      console.error('Failed to delete file', err);
      setDocumentsError(err?.message || 'Unable to delete file');
    }
  }, [business.id, refreshDocuments, loadDocumentTree]);

  const handleEmptyTrash = useCallback(async () => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    const confirmed = window.confirm('Empty all trash folders? This cannot be undone.');
    if (!confirmed) return;
    try {
      setDocumentsError('');
      setEmptyingTrash(true);
      await window.api?.emptyDocumentsTrash?.({ businessId: business.id });
      setMessage('Trash emptied');
      await refreshDocuments();
      await loadDocumentTree();
      setTimeout(() => setMessage(''), 2000);
    } catch (err) {
      console.error('Failed to empty trash', err);
      setDocumentsError(err?.message || 'Unable to empty trash');
    } finally {
      setEmptyingTrash(false);
    }
  }, [business.id, refreshDocuments, loadDocumentTree]);

  const handleRevealDocument = useCallback(async (filePath) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    setDocumentsError('');
    if (!filePath) {
      setDocumentsError('Document file not available');
      return;
    }
    try {
      const response = await window.api?.showItemInFolder?.(filePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to locate document on disk');
      }
    } catch (err) {
      console.error('Failed to reveal document', err);
      setDocumentsError(err?.message || 'Unable to locate document on disk');
    }
  }, []);

  const handleDeleteDocumentRecord = useCallback(async (doc) => {
    if (!DOCUMENT_FEATURES_ENABLED) {
      setDocumentsError('Document generation is disabled.');
      return;
    }
    if (!doc || doc.document_id == null) return;
    const title = doc.typeLabel
      ? `${doc.typeLabel}${doc.number ? ` #${doc.number}` : ''}`
      : 'this document';
    const confirmDelete = window.confirm(`Delete ${title}? This will remove it from the documents list.`);
    if (!confirmDelete) return;

    let removeFile = false;
    if (doc.file_path) {
      removeFile = window.confirm('Also remove the generated file from disk?');
    }

    try {
      setError('');
      if (doc.is_locked) {
        const unlock = window.confirm('This document is locked. Unlock and delete it now?');
        if (!unlock) return;
        try { await window.api?.setDocumentLock?.(doc.document_id, false); } catch (_) {}
      }
      await window.api?.deleteDocument?.(doc.document_id, { removeFile });
      setMessage('Document deleted');
      await refreshDocuments();
      await loadDocumentTree();
      setSelectedDocuments(prev => {
        const next = new Set(prev);
        next.delete(doc.document_id);
        return next;
      });
      window.api?.notifyJobsheetChange?.({
        type: 'documents-updated',
        businessId: business.id,
        jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
        documentId: doc.document_id
      });
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to delete document', err);
      setError(err?.message || 'Unable to delete document');
    }
  }, [refreshDocuments, loadDocumentTree, business.id]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      const match = WORKSPACE_SECTIONS.find(section => section.key === workspaceSection);
      const value = match ? match.key : 'jobsheets';
      window.localStorage.setItem(WORKSPACE_SECTION_STORAGE_KEY, value);
    } catch (err) {
      console.warn('Unable to persist workspace section', err);
    }
  }, [workspaceSection]);

  useEffect(() => {
    setError('');
    refreshJobsheets();
  }, [refreshJobsheets]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      window.localStorage.setItem(`ui:${business.id}:showArchived`, showArchived ? '1' : '0');
    } catch (_) {}
  }, [business.id, showArchived]);

  const handleToggleShowArchived = useCallback(() => {
    setShowArchived(prev => !prev);
  }, []);

  const handleArchiveToggle = useCallback(async (jobsheetId, archived) => {
    if (!window.api || typeof window.api.setJobsheetArchived !== 'function') {
      setError('Archive action unavailable');
      return;
    }
    try {
      await window.api.setJobsheetArchived(jobsheetId, archived);
      setMessage(archived ? 'Jobsheet archived' : 'Jobsheet unarchived');
      await refreshJobsheets();
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to toggle archive', err);
      setError(err?.message || 'Unable to update archive state');
    }
  }, [refreshJobsheets]);

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return;
    refreshDocuments();
  }, [refreshDocuments]);

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return;
    if (workspaceSection === 'documents') {
      loadDocumentTree();
    }
  }, [workspaceSection, loadDocumentTree]);

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return;
    if (typeof window === 'undefined') return;
    try {
      window.localStorage.setItem(DOCUMENT_COLUMNS_STORAGE_KEY, JSON.stringify(documentColumnsState));
    } catch (err) {
      console.warn('Unable to persist document columns preference', err);
    }
  }, [documentColumnsState]);

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return undefined;
    if (!columnsMenuOpen) return undefined;
    const handleClick = (event) => {
      if (columnsMenuRef.current && !columnsMenuRef.current.contains(event.target)) {
        setColumnsMenuOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [columnsMenuOpen]);

  useLayoutEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED) return;
    if (!columnsMenuOpen) return;
    const buttonEl = columnsMenuRef.current;
    const menuEl = columnsMenuContentRef.current;
    if (!buttonEl || !menuEl) return;

    const buttonRect = buttonEl.getBoundingClientRect();
    const menuHeight = menuEl.offsetHeight;
    const spaceBelow = window.innerHeight - buttonRect.bottom;
    const spaceAbove = buttonRect.top;

    if (spaceBelow < menuHeight + 12 && spaceAbove > spaceBelow) {
      setColumnsMenuAbove(true);
    } else {
      setColumnsMenuAbove(false);
    }
  }, [columnsMenuOpen, documentColumnsState, activeDocumentColumns]);

  useEffect(() => {
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload || payload.businessId !== business.id) return;
      if (payload.type === 'document-lock-toggled' && payload.documentId != null && typeof payload.locked === 'boolean') {
        const docId = Number(payload.documentId);
        setDocuments(prev => prev.map(d => (
          d && d.document_id === docId ? { ...d, is_locked: payload.locked ? 1 : 0 } : d
        )));
        return;
      }
      if (payload.type === 'documents-updated') {
        refreshDocuments();
        loadDocumentTree();
        const payloadJobsheetId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
        if (payloadJobsheetId != null) {
          setInlineEditorTargetId(payloadJobsheetId);
          setActiveJobsheetId(payloadJobsheetId);
          setInlineEditorVisible(true);
        }
        return;
      }
      if (payload.type === 'jobsheet-editor-focus') {
        const focusedId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
        if (payload.active) {
          setActiveJobsheetId(focusedId);
        } else if (focusedId != null) {
          setActiveJobsheetId(prev => (prev === focusedId ? null : prev));
        } else if (!focusedId) {
          setActiveJobsheetId(null);
        }
        return;
      }
      if (payload.type === 'jobsheet-deleted' && payload.jobsheetId != null) {
        const deletedId = Number(payload.jobsheetId);
        setActiveJobsheetId(prev => (prev === deletedId ? null : prev));
        if (inlineEditorTargetId != null && deletedId === inlineEditorTargetId) {
          setInlineEditorVisible(false);
          setInlineEditorTargetId(null);
        }
        return;
      }
      if (payload.type === 'jobsheet-created' && payload.jobsheetId != null) {
        const createdId = Number(payload.jobsheetId);
        if (inlineEditorVisible && inlineEditorTargetId == null) {
          setInlineEditorTargetId(createdId);
          setActiveJobsheetId(createdId);
        }
        if (payload.snapshot) {
          mergeJobsheetSnapshot(payload.snapshot);
        } else {
          refreshJobsheets();
        }
        return;
      }
      if (payload.type === 'jobsheet-load-request') {
        const requestedId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
        setActiveJobsheetId(requestedId);
        return;
      }
      if (payload.type === 'jobsheet-updated' && payload.snapshot) {
        if (inlineEditorVisible && inlineEditorTargetId == null && payload.snapshot.jobsheet_id != null) {
          const snapshotId = Number(payload.snapshot.jobsheet_id);
          setInlineEditorTargetId(snapshotId);
          setActiveJobsheetId(snapshotId);
        }
        mergeJobsheetSnapshot(payload.snapshot);
      } else {
        refreshJobsheets();
      }
    });
    return () => unsubscribe?.();
  }, [business.id, refreshJobsheets, refreshDocuments, loadDocumentTree, mergeJobsheetSnapshot, inlineEditorTargetId, inlineEditorVisible]);

  const handleChangeDocumentsFolder = useCallback(async () => {
    const api = window.api;
    if (!api || !api.updateBusinessSettings) {
      setError('Unable to update documents folder: API unavailable');
      return;
    }

    try {
      setError('');
      const previousPath = business.save_path || '';
      let selectedPath = null;
      if (typeof api.chooseDirectory === 'function') {
        selectedPath = await api.chooseDirectory({
          title: `Choose documents folder for ${business.business_name}`,
          defaultPath: business.save_path || undefined
        });
      } else {
        selectedPath = window.prompt('Enter documents folder path', business.save_path || '');
      }

      if (!selectedPath) return;
      if (typeof selectedPath === 'string') {
        selectedPath = selectedPath.trim();
      }
      if (!selectedPath) return;

      setUpdatingSavePath(true);
      const result = await api.updateBusinessSettings(business.id, { save_path: selectedPath });
      const updated = result?.record || { ...business, save_path: selectedPath };

      let relocationSummary = null;
      let relocationFailed = false;
      if ((previousPath || '') !== selectedPath && typeof api.relocateBusinessDocuments === 'function') {
        try {
          relocationSummary = await api.relocateBusinessDocuments({
            businessId: business.id,
            sourcePath: previousPath || undefined,
            targetPath: selectedPath
          });
        } catch (relocationError) {
          relocationFailed = true;
          console.error('Failed to relocate documents', relocationError);
          setError(relocationError?.message || 'Unable to move existing documents');
        }
      }

      onBusinessUpdate?.(updated);

      if (relocationSummary) {
        const movedCount = relocationSummary.moved?.length || 0;
        const skippedCount = relocationSummary.skipped?.length || 0;
        const errorCount = relocationSummary.errors?.length || 0;
        const summaryParts = [`moved ${movedCount}`];
        if (skippedCount) summaryParts.push(`skipped ${skippedCount}`);
        if (errorCount) summaryParts.push(`errors ${errorCount}`);
        setMessage(`Documents folder updated (${summaryParts.join(', ')})`);
        if (errorCount) {
          setError(`Unable to move ${errorCount} document${errorCount === 1 ? '' : 's'}. Check the folder and try again.`);
        }
      } else if (relocationFailed) {
        setMessage('Documents folder updated. Existing files were not moved.');
      } else {
        setMessage('Documents folder updated');
      }
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to update documents folder', err);
      setError(err?.message || 'Unable to update documents folder');
    } finally {
      setUpdatingSavePath(false);
    }
  }, [business, onBusinessUpdate]);

  const handleDeleteSelected = useCallback(async () => {
    if (!selectedDocuments.size) return;
    const ids = Array.from(selectedDocuments);
    const lockedIds = (normalizedDocuments || [])
      .filter(doc => ids.includes(doc.document_id) && doc.is_locked)
      .map(doc => doc.document_id);
    const confirmMessage = ids.length === 1
      ? 'Delete the selected document?'
      : `Delete ${ids.length} selected documents?`;
    if (!window.confirm(confirmMessage)) return;

    const hasFiles = normalizedDocuments.some(doc => ids.includes(doc.document_id) && doc.fileAvailable);
    let removeFiles = false;
    if (hasFiles) {
      removeFiles = window.confirm('Also remove the generated files from disk?');
    }

    try {
      setError('');
      if (lockedIds.length) {
        const doUnlock = window.confirm(`${lockedIds.length} selected document(s) are locked. Unlock and delete all?`);
        if (!doUnlock) return;
        await Promise.all(lockedIds.map(id => window.api?.setDocumentLock?.(id, false)));
      }
      await Promise.all(ids.map(id => window.api?.deleteDocument?.(id, { removeFile: removeFiles })));
      setMessage(ids.length === 1 ? 'Document deleted' : 'Selected documents deleted');
      await refreshDocuments();
      setSelectedDocuments(new Set());
      const impactedJobsheets = new Set();
      normalizedDocuments.forEach(doc => {
        if (ids.includes(doc.document_id) && doc.jobsheet_id != null) {
          impactedJobsheets.add(Number(doc.jobsheet_id));
        }
      });
      if (impactedJobsheets.size) {
        impactedJobsheets.forEach(id => {
          window.api?.notifyJobsheetChange?.({
            type: 'documents-updated',
            businessId: business.id,
            jobsheetId: id,
            documentIds: ids
          });
        });
      } else {
        window.api?.notifyJobsheetChange?.({
          type: 'documents-updated',
          businessId: business.id,
          documentIds: ids
        });
      }
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to delete selected documents', err);
      setError(err?.message || 'Unable to delete selected documents');
    }
  }, [selectedDocuments, normalizedDocuments, refreshDocuments, business.id]);

  const handleUnlockSelected = useCallback(async () => {
    if (!selectedDocuments.size) return;
    const ids = Array.from(selectedDocuments);
    const lockedIds = (normalizedDocuments || [])
      .filter(doc => ids.includes(doc.document_id) && doc.is_locked)
      .map(doc => doc.document_id);
    if (!lockedIds.length) return;
    const confirmMessage = lockedIds.length === 1
      ? 'Unlock the selected document?'
      : `Unlock ${lockedIds.length} selected documents?`;
    if (!window.confirm(confirmMessage)) return;
    try {
      await Promise.all(lockedIds.map(id => window.api?.setDocumentLock?.(id, false)));
      setMessage(lockedIds.length === 1 ? 'Document unlocked' : 'Selected documents unlocked');
      await refreshDocuments();
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to unlock selected documents', err);
      setError(err?.message || 'Unable to unlock selected documents');
    }
  }, [selectedDocuments, normalizedDocuments, refreshDocuments]);

  const normalizedDocuments = useMemo(() => {
    return (documents || []).map(doc => {
      const typeLabel = DOCUMENT_TYPE_LABELS[doc.doc_type] || startCaseKey(doc.doc_type || 'document');
      const displayClient = doc.display_client_name || doc.client_name || doc.joined_client_name || '';
      const displayEvent = doc.display_event_name || doc.event_name || doc.joined_event_name || '';
      const eventDateRaw = doc.display_event_date || doc.joined_event_date || doc.event_date || '';
      const documentDateRaw = doc.document_date || '';
      const eventDateIso = eventDateRaw ? formatDateInput(eventDateRaw) : '';
      const formattedEventDate = eventDateIso ? formatCompactDate(eventDateIso) : '—';
      const formattedDocumentDate = documentDateRaw ? formatCompactDate(documentDateRaw) : '';
      const createdAtDisplay = formatCompactDate(doc.created_at);
      const createdAtFull = doc.created_at ? formatTimestampDisplay(doc.created_at) : '';
      const statusLabel = (doc.status || 'draft').replace(/_/g, ' ');
      const fileName = doc.file_name || (doc.file_path ? doc.file_path.split(/[\\/]+/).filter(Boolean).pop() : '');
      const displayLabel = doc.display_label || doc.definition_label || typeLabel;
      const filePrefix = '';
      const fileSuffix = '';
      const folderPath = doc.folder_path || '';
      const fileAvailable = doc.file_available !== false && Boolean(doc.file_path);

      return {
        ...doc,
        typeLabel,
        displayClient: displayClient || '—',
        displayEvent: displayEvent || '',
        eventDateIso,
        formattedEventDate,
        formattedDocumentDate,
        createdAtDisplay,
        createdAtFull,
        statusLabel,
        fileName,
        displayLabel,
        filePrefix,
        fileSuffix,
        folderPath,
        fileAvailable
      };
    });
  }, [documents]);

  // Helpers for matching PDFs to workbooks by base filename (no extension)
  const baseNameNoExt = useCallback((fp) => {
    const name = fp ? String(fp).split(/[\\/]+/).pop() : '';
    return name ? name.replace(/\.[^.]+$/, '') : '';
  }, []);

  const pdfBaseNames = useMemo(() => {
    const set = new Set();
    (normalizedDocuments || []).forEach(doc => {
      const path = doc?.file_path || '';
      if (path && path.toLowerCase().endsWith('.pdf')) {
        const base = baseNameNoExt(path);
        if (base) set.add(base);
      }
    });
    return set;
  }, [normalizedDocuments, baseNameNoExt]);

  useEffect(() => {
    setSelectedDocuments(prev => {
      const next = new Set();
      normalizedDocuments.forEach(doc => {
        if (prev.has(doc.document_id)) {
          next.add(doc.document_id);
        }
      });
      return next;
    });
  }, [normalizedDocuments]);

  const toggleDocumentSelection = useCallback((docId, checked) => {
    setSelectedDocuments(prev => {
      const next = new Set(prev);
      if (checked) {
        next.add(docId);
      } else {
        next.delete(docId);
      }
      return next;
    });
  }, []);

  const handleSelectGroupDocs = useCallback((docIds, checked) => {
    setSelectedDocuments(prev => {
      const next = new Set(prev);
      docIds.forEach(id => {
        if (checked) {
          next.add(id);
        } else {
          next.delete(id);
        }
      });
      return next;
    });
  }, []);

  const handleToggleColumn = useCallback((columnKey) => {
    setDocumentColumnsState(prev => {
      if (DOCUMENT_COLUMNS.find(column => column.key === columnKey)?.always) {
        return prev;
      }
      const current = prev?.[columnKey] !== false;
      return {
        ...prev,
        [columnKey]: !current
      };
    });
  }, []);

  const selectedCount = selectedDocuments.size;
  const documentsSearchValue = documentsSearch.trim().toLowerCase();

  const filteredDocuments = useMemo(() => {
    if (!documentsSearchValue) return normalizedDocuments;
    return normalizedDocuments.filter(doc => {
      const haystack = [
        doc.typeLabel,
        doc.displayLabel,
        doc.displayClient,
        doc.displayEvent,
        doc.statusLabel,
        doc.formattedEventDate,
        doc.formattedDocumentDate,
        doc.createdAtDisplay,
        doc.createdAtFull,
        doc.doc_type,
        doc.file_path,
        doc.fileName,
        doc.filePrefix,
        doc.folderPath,
        doc.number ? `#${doc.number}` : '',
        doc.document_id != null ? String(doc.document_id) : ''
      ].join(' ').toLowerCase();
      return haystack.includes(documentsSearchValue);
    });
  }, [normalizedDocuments, documentsSearchValue]);

  const groupedDocuments = useMemo(() => {
    if (documentsGroup === 'none') {
      return [];
    }

    const groups = new Map();
    const ensureGroup = (key, label) => {
      const mapKey = key || '__missing__';
      if (!groups.has(mapKey)) {
        groups.set(mapKey, { key: mapKey, label: label || 'Other', items: [] });
      }
      return groups.get(mapKey);
    };

    filteredDocuments.forEach(doc => {
      if (documentsGroup === 'doc_type') {
        const key = doc.doc_type || 'unknown';
        const entry = ensureGroup(key, doc.typeLabel || 'Other');
        entry.items.push(doc);
      } else if (documentsGroup === 'client') {
        const key = (doc.displayClient && doc.displayClient !== '—') ? doc.displayClient : 'No client';
        const entry = ensureGroup(key, key);
        entry.items.push(doc);
      } else if (documentsGroup === 'event_date') {
        const key = doc.eventDateIso || 'no-date';
        const label = doc.eventDateIso ? doc.formattedEventDate : 'No event date';
        const entry = ensureGroup(key, label);
        entry.items.push(doc);
      }
    });

    const result = Array.from(groups.values());
    if (documentsGroup === 'event_date') {
      result.sort((a, b) => {
        if (a.key === 'no-date') return 1;
        if (b.key === 'no-date') return -1;
        return a.key.localeCompare(b.key);
      });
    } else {
      result.sort((a, b) => a.label.localeCompare(b.label, 'en', { sensitivity: 'base' }));
    }

    return result;
  }, [documentsGroup, filteredDocuments]);

  const activeDocumentColumns = useMemo(() => (
    DOCUMENT_COLUMNS.filter(column => column.always || documentColumnsState[column.key] !== false)
  ), [documentColumnsState]);

  const canDeleteSelected = selectedCount > 0 && !documentsLoading;
  const hasLockedSelected = useMemo(() => {
    if (!selectedCount) return false;
    const ids = new Set(Array.from(selectedDocuments));
    return (normalizedDocuments || []).some(doc => ids.has(doc.document_id) && doc.is_locked);
  }, [selectedCount, selectedDocuments, normalizedDocuments]);
  const canUnlockSelected = selectedCount > 0 && hasLockedSelected && !documentsLoading;

  const documentsGroupLabel = DOCUMENT_GROUP_OPTIONS.find(option => option.value === documentsGroup)?.label || 'All Documents';
  const headerSubtitle = documentsGroup === 'none'
    ? `${filteredDocuments.length} item${filteredDocuments.length === 1 ? '' : 's'}`
    : `${filteredDocuments.length} items · ${documentsGroupLabel}`;

  const emptyStateMessage = documentsSearchValue
    ? 'No documents match your search.'
    : documentsGroup === 'none'
      ? 'No documents generated yet.'
      : 'No documents available in this group yet.';

  const renderDocumentTable = useCallback((items) => {
    if (!items.length) return null;

    const docIds = items
      .map(doc => doc.document_id)
      .filter(id => id != null);
    const allSelected = docIds.length > 0 && docIds.every(id => selectedDocuments.has(id));
    const someSelected = docIds.some(id => selectedDocuments.has(id));

    return (
      <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white shadow-sm">
        <table className="w-full table-auto text-sm">
          <thead className="bg-slate-50 text-xs font-semibold uppercase tracking-wide text-slate-600">
            <tr>
              <th className="w-12 px-3 py-3 text-left">
                <IndeterminateCheckbox
                  checked={allSelected}
                  indeterminate={!allSelected && someSelected}
                  onChange={event => handleSelectGroupDocs(docIds, event.target.checked)}
                  aria-label="Select group"
                />
              </th>
              {activeDocumentColumns.map(column => {
                const alignClass = column.align === 'right'
                  ? 'text-right'
                  : column.align === 'center'
                    ? 'text-center'
                    : 'text-left';
                return (
                  <th
                    key={column.key}
                    className={`px-3 py-3 ${alignClass}`}
                  >
                    {column.label}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {items.map((doc, index) => {
              const rowSelected = selectedDocuments.has(doc.document_id);
              const rowClass = rowSelected
                ? 'bg-indigo-50/80'
                : index % 2 === 0
                  ? 'bg-white'
                  : 'bg-slate-50';
              const typeBadge = doc.typeLabel + (doc.number ? ` #${doc.number}` : '');
              const primaryText = doc.fileName || doc.displayLabel || typeBadge;
              const secondaryTexts = [];
              if (doc.displayLabel && doc.displayLabel !== primaryText) secondaryTexts.push(doc.displayLabel);
              if (doc.filePrefix) secondaryTexts.push(doc.filePrefix);
              if (typeBadge && typeBadge !== primaryText && typeBadge !== doc.displayLabel) secondaryTexts.push(typeBadge);
              const tooltipText = doc.file_path || doc.folderPath || primaryText;
              return (
                <tr key={doc.document_id} className={`transition ${rowClass}`}>
                  <td className="align-top px-3 py-3">
                    <input
                      type="checkbox"
                      className="mt-1 h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                      checked={rowSelected}
                      onChange={event => toggleDocumentSelection(doc.document_id, event.target.checked)}
                      aria-label="Select document"
                    />
                  </td>
                  {activeDocumentColumns.map(column => {
                    const alignClass = column.align === 'right'
                      ? 'text-right'
                      : column.align === 'center'
                        ? 'text-center'
                        : 'text-left';
                    let cell = null;
                    if (column.key === 'document') {
                      cell = (
                        <div className="flex items-center gap-3">
                          <span className="text-lg" role="img" aria-label={doc.typeLabel}>{getDocumentIcon(doc.doc_type)}</span>
                          <div className="min-w-0 space-y-1">
                            <div
                              className="text-sm font-medium text-slate-700 truncate"
                              title={tooltipText}
                            >
                              {primaryText}
                            </div>
                            {secondaryTexts.map((text, idx) => (
                              <div key={`${text}-${idx}`} className="text-xs text-slate-500 truncate" title={text}>{text}</div>
                            ))}
                            {doc.folderPath ? (
                              <div className="text-[11px] text-slate-400 truncate" title={doc.folderPath}>{doc.folderPath}</div>
                            ) : null}
                          </div>
                          {!doc.fileAvailable ? (
                            <div className="text-xs font-medium text-amber-600">Missing</div>
                          ) : null}
                        </div>
                      );
                    } else if (column.key === 'client') {
                      cell = (
                        <div className="text-slate-600">
                          <div className="font-medium" title={doc.displayClient}>{doc.displayClient}</div>
                          {doc.displayEvent ? (
                            <div className="text-xs text-slate-500" title={doc.displayEvent}>{doc.displayEvent}</div>
                          ) : null}
                        </div>
                      );
                    } else if (column.key === 'event_date') {
                      cell = (
                        <div className="text-slate-600">
                          <span title={doc.eventDateIso || undefined}>{doc.formattedEventDate}</span>
                          {doc.formattedDocumentDate && doc.formattedDocumentDate !== doc.formattedEventDate ? (
                            <div className="text-xs text-slate-500">Doc: {doc.formattedDocumentDate}</div>
                          ) : null}
                        </div>
                      );
                    } else if (column.key === 'created') {
                      cell = (
                        <div className="text-slate-600">
                          <span title={doc.createdAtFull || undefined}>{doc.createdAtDisplay}</span>
                        </div>
                      );
                    } else if (column.key === 'amount') {
                      cell = (
                        <div className="font-semibold text-slate-700">
                          {toCurrency(doc.total_amount)}
                        </div>
                      );
                    } else if (column.key === 'actions') {
                      const isLocked = Boolean(doc?.is_locked);
                      const isWorkbook = (doc?.doc_type || '').toLowerCase() === 'workbook';
                      const fileExists = doc?.fileAvailable !== false && Boolean(doc?.file_path);
                      const workbookHasPdf = isWorkbook && fileExists ? pdfBaseNames.has(baseNameNoExt(doc.file_path)) : false;
                      cell = (
                        <div className="flex flex-wrap justify-end gap-1.5">
                          <IconButton
                            label={isLocked ? 'Unlock' : 'Lock'}
                            onClick={async () => {
                              try {
                                await window.api?.setDocumentLock?.(doc.document_id, !isLocked);
                                await refreshDocuments();
                              } catch (err) {
                                console.error('Failed to toggle document lock', err);
                                setError(err?.message || 'Unable to toggle document lock');
                              }
                            }}
                            disabled={!doc?.document_id}
                    className={isLocked ? 'border-red-300 text-red-600 hover:bg-red-50' : 'border-green-300 text-green-600 hover:bg-green-50'}
                          >
                            <span className="text-base" aria-hidden>{isLocked ? '🔒' : '🔓'}</span>
                          </IconButton>
                          <IconButton
                            label="Open document"
                            onClick={() => handleOpenDocumentFile(doc.file_path)}
                            disabled={!fileExists}
                          >
                            <OpenIcon />
                          </IconButton>
                          <IconButton
                            label="Reveal document in Finder"
                            onClick={() => handleRevealDocument(doc.file_path)}
                            disabled={!fileExists}
                          >
                            <RevealIcon />
                          </IconButton>
                          {isWorkbook ? (
                            <button
                              type="button"
                              onClick={() => handleExportWorkbookPdf(doc)}
                              disabled={!fileExists || isLocked || workbookHasPdf}
                              className="inline-flex items-center rounded border border-indigo-200 px-2 py-1 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-60"
                            >
                              Export
                            </button>
                          ) : null}
                          <IconButton
                            label="Delete document"
                            onClick={() => handleDeleteDocumentRecord(doc)}
                            disabled={doc?.document_id == null}
                            className="border-red-200 text-red-600 hover:bg-red-50"
                          >
                            <DeleteIcon />
                          </IconButton>
                        </div>
                      );
                    }

                    return (
                      <td
                        key={column.key}
                        className={`align-top px-3 py-3 text-sm text-slate-600 ${alignClass}`}
                      >
                        {cell}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }, [activeDocumentColumns, handleDeleteDocumentRecord, handleOpenDocumentFile, handleRevealDocument, handleSelectGroupDocs, selectedDocuments, toggleDocumentSelection]);

  // Mirror inline documents pane visually (no generation), across all files
  const documentsContent = useMemo(() => {
    const excelDocs = filteredDocuments.filter(doc => (doc?.file_path || '').toLowerCase().endsWith('.xlsx'));
    const pdfDocs = filteredDocuments.filter(doc => (doc?.file_path || '').toLowerCase().endsWith('.pdf'));

    // No export-all in main pane

    const ExcelPane = (
      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <div className="text-sm font-semibold text-slate-700">Excel</div>
          <div className="flex items-center gap-2 text-xs">
            <button type="button" onClick={handleRefreshDocuments} className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50">Refresh</button>
          </div>
        </div>
        <div className="rounded border border-slate-200 bg-white p-2 space-y-1">
          {excelDocs.map(doc => {
            const label = doc.fileName || doc.displayLabel || 'Workbook';
            const locked = Boolean(doc?.is_locked);
            const generated = Boolean(doc?.fileAvailable);
            const hasPdf = doc?.fileAvailable ? pdfBaseNames.has(baseNameNoExt(doc.file_path)) : false;
            return (
              <div key={`xl:${doc.document_id || doc.file_path}`} className="flex items-center justify-between rounded px-2 py-2">
                <div className="min-w-0">
                  <div className={`flex items-center gap-2 text-sm font-medium truncate ${generated ? 'text-slate-700' : 'text-slate-300 opacity-70'}`}>
                    <span aria-hidden>{generated ? '✅' : '❌'}</span>
                    <span className="truncate" title={doc.file_path}>{label}</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <IconButton
                    label={locked ? 'Unlock workbook' : 'Lock workbook'}
                    onClick={async () => {
                      try {
                        await window.api?.setDocumentLock?.(doc.document_id, !locked);
                        // Notify without triggering full refresh in main pane
                        window.api?.notifyJobsheetChange?.({
                          type: 'document-lock-toggled',
                          businessId: business.id,
                          jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
                          documentId: doc.document_id,
                          locked: !locked
                        });
                      } catch (err) {
                        console.error('Failed to toggle lock', err);
                        setError(err?.message || 'Unable to toggle lock');
                      }
                    }}
                    disabled={!generated || !doc?.document_id}
                    className={locked ? 'border-red-300 text-red-600 hover:bg-red-50' : 'border-green-300 text-green-600 hover:bg-green-50'}
                  >
                    <span className="text-base" aria-hidden>{locked ? '🔒' : '🔓'}</span>
                  </IconButton>
                  {/* Export removed in main pane */}
                  <div className="flex items-center gap-1.5">
                    <IconButton label="Open" onClick={() => handleOpenDocumentFile(doc.file_path)} disabled={!generated}><OpenIcon className="h-3.5 w-3.5" /></IconButton>
                    <IconButton label="Reveal in Finder" onClick={() => handleRevealDocument(doc.file_path)} disabled={!generated}><RevealIcon className="h-3.5 w-3.5" /></IconButton>
                    <IconButton label="Delete" onClick={() => handleDeleteDocumentRecord(doc)} disabled={!generated || doc?.document_id == null} className="border-red-200 text-red-600 hover:bg-red-50"><DeleteIcon className="h-3.5 w-3.5" /></IconButton>
                  </div>
                </div>
              </div>
            );
          })}
          {excelDocs.length === 0 ? (
            <div className="px-2 py-2 text-sm text-slate-500">No Excel workbooks.</div>
          ) : null}
        </div>
      </div>
    );

    const PdfPane = (
      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <div className="text-sm font-semibold text-slate-700">PDFs</div>
          <div className="flex items-center gap-2 text-xs">
            <button type="button" onClick={handleRefreshDocuments} className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50">Refresh</button>
          </div>
        </div>
        <div className="rounded border border-slate-200 bg-white p-2 space-y-1">
          {pdfDocs.map(doc => {
            const label = doc.displayLabel || doc.fileName || 'PDF';
            const exported = Boolean(doc?.fileAvailable);
            const locked = Boolean(doc?.is_locked);
            return (
              <div key={`pdf:${doc.document_id || doc.file_path}`} className="flex items-center justify-between rounded px-2 py-2">
                <div className="min-w-0">
                  <div className={`flex items-center gap-2 text-sm font-medium truncate ${exported ? 'text-slate-700' : 'text-slate-300 opacity-70'}`}>
                    <span aria-hidden>{exported ? '✅' : '❌'}</span>
                    <span className="truncate" title={doc.file_path}>{label}</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <IconButton label={locked ? 'Unlock PDF' : 'Lock PDF'} onClick={async () => {
                    try {
                      await window.api?.setDocumentLock?.(doc.document_id, !locked);
                      window.api?.notifyJobsheetChange?.({
                        type: 'document-lock-toggled',
                        businessId: business.id,
                        jobsheetId: doc.jobsheet_id != null ? Number(doc.jobsheet_id) : null,
                        documentId: doc.document_id,
                        locked: !locked
                      });
                    } catch (err) { console.error('Failed to toggle lock', err); setError(err?.message || 'Unable to toggle lock'); }
                  }} disabled={!exported || !doc?.document_id} className={locked ? 'border-red-300 text-red-600 hover:bg-red-50' : 'border-green-300 text-green-600 hover:bg-green-50'}><span className="text-base" aria-hidden>{locked ? '🔒' : '🔓'}</span></IconButton>
                  <div className="flex items-center gap-1.5">
                    <IconButton label="Open" onClick={() => handleOpenDocumentFile(doc.file_path)} disabled={!exported}><OpenIcon className="h-3.5 w-3.5" /></IconButton>
                    <IconButton label="Reveal in Finder" onClick={() => handleRevealDocument(doc.file_path)} disabled={!exported}><RevealIcon className="h-3.5 w-3.5" /></IconButton>
                    <IconButton label="Delete" onClick={() => handleDeleteDocumentRecord(doc)} disabled={!exported || doc?.document_id == null} className="border-red-200 text-red-600 hover:bg-red-50"><DeleteIcon className="h-3.5 w-3.5" /></IconButton>
                  </div>
                </div>
              </div>
            );
          })}
          {pdfDocs.length === 0 ? (
            <div className="px-2 py-2 text-sm text-slate-500">No PDFs.</div>
          ) : null}
        </div>
      </div>
    );

    return (
      <div className="space-y-4">
        {ExcelPane}
        {PdfPane}
      </div>
    );
  }, [filteredDocuments, handleRefreshDocuments, handleOpenDocumentFile, handleRevealDocument, handleDeleteDocumentRecord, pdfBaseNames, baseNameNoExt, setError, refreshDocuments]);

  const documentTreeRoot = documentTree?.root || null;
  const documentTreeTrash = documentTree?.trash || null;
  const documentTreePath = documentTree?.rootPath || business.save_path || '';
  const documentsConfigured = Boolean((business.save_path || '').trim());

  const openJobsheetWindow = useCallback((jobsheetId) => {
    const api = window.api;
    if (!api || !api.openJobsheetWindow) {
      setError('Unable to open editor window: API unavailable');
      return;
    }
    api.openJobsheetWindow({
      businessId: business.id,
      businessName: business.business_name,
      jobsheetId
    });
  }, [business.id, business.business_name]);

  const scrollInlineEditorIntoView = useCallback(() => {
    try {
      const anchor = document.getElementById('inline-jobsheet-editor');
      if (!anchor) return;
      // If user is already typing inside the editor, avoid interrupting focus
      const ae = document.activeElement;
      if (ae && anchor.contains(ae) && (/^(input|textarea|select)$/i).test(ae.tagName)) {
        return;
      }
      // If the editor is already within view (account for sticky header), skip scrolling
      const sticky = document.getElementById('jobsheet-sticky-header');
      const stickyHeight = sticky ? (sticky.getBoundingClientRect().height || 0) : 120;
      const viewTop = stickyHeight + 8;
      const viewBottom = window.innerHeight || document.documentElement.clientHeight || 0;
      const rect = anchor.getBoundingClientRect();
      const topAfterHeader = rect.top - viewTop;
      const isMostlyVisible = topAfterHeader >= -80 && rect.top < viewBottom * 0.75;
      if (isMostlyVisible) return;
      const extraGap = 12;
      const top = anchor.getBoundingClientRect().top + window.scrollY - (stickyHeight + extraGap);
      try {
        window.scrollTo({ top: Math.max(top, 0), behavior: 'smooth' });
      } catch (_err) {
        window.scrollTo(0, Math.max(top, 0));
      }
    } catch (_err) {}
  }, []);

  const handleNew = useCallback(() => {
    setActiveJobsheetId(null);
    setInlineEditorTargetId(null);
    setInlineEditorVisible(true);
    setInlineEditorSession(prev => prev + 1);
    // Scroll to inline editor after it mounts
    setTimeout(() => scrollInlineEditorIntoView(), 250);
  }, [scrollInlineEditorIntoView]);

  const handleOpenExisting = useCallback((jobsheetId) => {
    if (!jobsheetId) return;
    const numericId = Number(jobsheetId);
    setActiveJobsheetId(numericId);
    setInlineEditorTargetId(numericId);
    setInlineEditorVisible(true);
    setInlineEditorSession(prev => (numericId !== inlineEditorTargetId ? prev + 1 : prev));
    // Scroll to inline editor after it mounts
    setTimeout(() => scrollInlineEditorIntoView(), 250);
  }, [inlineEditorTargetId, scrollInlineEditorIntoView]);

  const handleDelete = useCallback(async (jobsheetId) => {
    if (!jobsheetId) return;
    const confirmed = window.confirm('Delete this jobsheet? This cannot be undone.');
    if (!confirmed) return;
    setDeletingId(jobsheetId);
    setError('');
    try {
      const api = window.api;
      if (!api || !api.deleteAhmenJobsheet) {
        setError('Unable to delete jobsheet: API unavailable');
        setDeletingId(null);
        return;
      }
      let cascaded = false;
      try {
        const deep = window.confirm('Also delete all related documents and emails (and move files to trash)? Click OK for full removal, Cancel for jobsheet only.');
        if (deep && api.deleteJobsheetCompletely) {
          cascaded = true;
          await api.deleteJobsheetCompletely({ businessId: business.id, jobsheetId, removeFiles: true });
        }
      } catch (_) {}
      if (!cascaded) {
        await api.deleteAhmenJobsheet(jobsheetId);
      }
      setMessage('Jobsheet deleted');
      await refreshJobsheets();
      window.api?.notifyJobsheetChange?.({ type: 'jobsheet-deleted', businessId: business.id, jobsheetId });
    } catch (err) {
      console.error('Failed to delete jobsheet', err);
      setError(err?.message || 'Unable to delete jobsheet');
    } finally {
      setDeletingId(null);
    }
  }, [refreshJobsheets, business.id]);

  const handleStatusChange = useCallback(async (jobsheetId, nextStatus) => {
    if (!jobsheetId || !nextStatus) return;
    const normalized = normalizeStatus(nextStatus) || 'enquiry';
    setStatusUpdatingId(jobsheetId);
    setError('');
    try {
      const api = window.api;
      if (!api || !api.updateAhmenJobsheetStatus) {
        setError('Unable to update status: API unavailable');
        return;
      }
      await api.updateAhmenJobsheetStatus(jobsheetId, normalized);
      setJobsheets(prev => prev.map(job => (
        job.jobsheet_id === jobsheetId
          ? normalizeJobsheet({ ...job, status: normalized })
          : job
      )));
      setMessage('Status updated');
      setTimeout(() => setMessage(''), 1500);
      window.api?.notifyJobsheetChange?.({
        type: 'jobsheet-updated',
        businessId: business.id,
        jobsheetId,
        snapshot: {
          jobsheet_id: jobsheetId,
          status: normalized
        }
      });
    } catch (err) {
      console.error('Failed to update jobsheet status', err);
      setError(err?.message || 'Unable to update status');
    } finally {
      setStatusUpdatingId(null);
    }
  }, [business.id, normalizeJobsheet]);

  const handleCloseInlineEditor = useCallback(() => {
    setInlineEditorVisible(false);
    activeJobsheetIdRef.current = null;
    setInlineEditorTargetId(null);
    setActiveJobsheetId(null);
  }, []);

  const handlePopoutEditor = useCallback(() => {
    openJobsheetWindow(inlineEditorTargetId ?? undefined);
    setInlineEditorVisible(false);
    setInlineEditorTargetId(null);
  }, [inlineEditorTargetId, openJobsheetWindow]);

  const inlineEditorKey = `jobsheet-editor-${inlineEditorSession}`;

  useEffect(() => {
    if (inlineEditorTargetId != null && !inlineEditorVisible) {
      setInlineEditorVisible(true);
    }
  }, [inlineEditorTargetId, inlineEditorVisible]);


  const handleSort = useCallback((columnKey) => {
    if (!columnKey) return;
    setSortConfig(prev => {
      if (prev.key === columnKey) {
        return { key: columnKey, direction: prev.direction === 'asc' ? 'desc' : 'asc' };
      }
      return { key: columnKey, direction: columnKey === 'client_name' ? 'asc' : 'desc' };
    });
  }, []);

  // Persist jobsheet sort when enabled
  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      if (window.localStorage.getItem('app:persistUiState') === 'true') {
        window.localStorage.setItem(`ui:${business.id}:jobsheetSort`, JSON.stringify(sortConfig || {}));
      }
    } catch (_err) {}
  }, [business.id, sortConfig]);

  const workspaceToasts = [];
  if (error) workspaceToasts.push({ id: 'workspace-error', tone: 'error', text: error });
  if (message) workspaceToasts.push({ id: 'workspace-message', tone: 'success', text: message });

  return (
    <div className="min-h-screen bg-slate-100">
      <ToastOverlay notices={workspaceToasts} />
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-screen-2xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold text-slate-800">{business.business_name}</h1>
            <p className="text-sm text-slate-500">Manage jobsheets, documents, and templates in one workspace.</p>
          </div>
          {/* Switch business removed */}
        </div>
      </header>

      <main className="max-w-screen-2xl mx-auto px-6 py-6 space-y-6">

        <div className="flex flex-col gap-6 lg:flex-row">
          <nav className="sticky top-4 z-30 flex-shrink-0 self-start md:w-56 lg:w-64">
            <div className="space-y-2" role="tablist" aria-orientation="vertical">
              {WORKSPACE_SECTIONS.map(section => {
                const isActive = workspaceSection === section.key;
                const icon = section.icon ?? getWorkspaceIcon(section.key);
                return (
                  <button
                    key={section.key}
                    type="button"
                    role="tab"
                    aria-selected={isActive}
                    onClick={() => setWorkspaceSection(section.key)}
                    className={`group flex w-full items-center gap-3 rounded-lg border px-3 py-3 text-left transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${isActive ? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold shadow-sm' : 'border-transparent bg-white text-slate-600 hover:bg-slate-50 hover:border-slate-200'}`}
                  >
                    <span className={`flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-full text-lg transition ${isActive ? 'bg-indigo-100 text-indigo-700 shadow-sm' : 'bg-slate-100 text-slate-500 group-hover:bg-slate-200 group-hover:text-slate-700'}`}>
                      {icon}
                    </span>
                    <span className="flex-1">
                      <span className="block text-sm font-semibold">{section.label}</span>
                      <span className="mt-1 block text-xs text-slate-500">{section.description}</span>
                    </span>
                  </button>
                );
              })}
            </div>
          </nav>

          <div className="flex-1 space-y-6">
            {workspaceSection === 'jobsheets' ? (
              <section className="space-y-4">
                <JobsheetList
                  business={business}
                  jobsheets={jobsheets}
                  onOpen={handleOpenExisting}
                  onNew={handleNew}
                  onDelete={handleDelete}
                  onStatusChange={handleStatusChange}
                  onArchiveToggle={handleArchiveToggle}
                  includeArchived={showArchived}
                  onToggleIncludeArchived={handleToggleShowArchived}
                  loading={listLoading}
                  deletingId={deletingId}
                  statusUpdatingId={statusUpdatingId}
                  sortConfig={sortConfig}
                  onSort={handleSort}
                  activeJobsheetId={activeJobsheetId}
                />
                <div id="inline-jobsheet-editor">
                  <InlineJobsheetEditorPanel
                  business={business}
                  visible={inlineEditorVisible}
                  jobsheetId={inlineEditorTargetId}
                  sessionKey={inlineEditorKey}
                  onClose={handleCloseInlineEditor}
                  onOpenInWindow={handlePopoutEditor}
                  />
                </div>
              </section>
            ) : null}

            {workspaceSection === 'documents' ? (
              <section className="space-y-4">
                <div className="grid gap-4 lg:grid-cols-[320px,1fr]">
                  <DocumentTreeView
                    root={documentTreeRoot}
                    trash={documentTreeTrash}
                    rootPath={documentTreePath}
                    loading={documentTreeLoading}
                    error={documentTreeError}
                    onRefresh={loadDocumentTree}
                    onOpen={handleOpenTreeNode}
                    onReveal={handleRevealTreeNode}
                    onDeleteFolder={handleDeleteTreeFolder}
                    onDeleteFile={handleDeleteTreeFile}
                    onEmptyTrash={handleEmptyTrash}
                    emptyingTrash={emptyingTrash}
                    isConfigured={documentsConfigured}
                    collapsed={documentTreeCollapsed}
                    onCollapsedChange={value => setDocumentTreeCollapsed(Boolean(value))}
                    persist={persistUi}
                    persistKey={`ui:${business.id}:documents`}
                  />
                  <div className="space-y-4">
                    <div className="rounded-lg border border-slate-200 bg-white p-4 space-y-4">
                      <div className="flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
                        <div>
                          <h2 className="text-lg font-semibold text-slate-700">Documents</h2>
                          <p className="text-sm text-slate-500">
                            {headerSubtitle}
                            <span className="ml-2 inline-block align-middle text-xs text-slate-400 w-[64px]">
                              {showDocumentsLoading ? 'Loading…' : '\u00A0'}
                            </span>
                          </p>
                        </div>
                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            type="button"
                            onClick={handleRefreshDocuments}
                            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                          >
                            Refresh list
                          </button>
                          <button
                            type="button"
                            onClick={handleOpenDocumentsFolder}
                            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                          >
                            Open folder
                          </button>
                          <button
                            type="button"
                            onClick={handleUnlockSelected}
                            disabled={!canUnlockSelected}
                            className="inline-flex items-center rounded border border-green-200 px-3 py-1.5 text-xs font-medium text-green-700 hover:bg-green-50 disabled:cursor-not-allowed disabled:opacity-60"
                          >
                            Unlock selected
                          </button>
                          <button
                            type="button"
                            onClick={handleDeleteSelected}
                            disabled={!canDeleteSelected}
                            className="inline-flex items-center rounded border border-red-200 px-3 py-1.5 text-xs font-medium text-red-600 hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-60"
                          >
                            Delete selected
                          </button>
                        </div>
                      </div>
                      <div className="grid gap-3 md:grid-cols-[minmax(0,1fr),200px,auto]">
                        <div className="flex items-center gap-2">
                          <label className="sr-only" htmlFor="documents-search">Search documents</label>
                          <input
                            id="documents-search"
                            type="search"
                            value={documentsSearch}
                            onChange={event => setDocumentsSearch(event.target.value)}
                            placeholder="Search documents"
                            className="w-full rounded border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                          />
                          {documentsSearchValue ? (
                            <button
                              type="button"
                              onClick={() => setDocumentsSearch('')}
                              className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100"
                            >
                              Clear
                            </button>
                          ) : null}
                        </div>
                        <div>
                          <label className="sr-only" htmlFor="documents-group">Group documents</label>
                          <select
                            id="documents-group"
                            value={documentsGroup}
                            onChange={event => setDocumentsGroup(event.target.value)}
                            className="w-full rounded border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                          >
                            {DOCUMENT_GROUP_OPTIONS.map(option => (
                              <option key={option.value} value={option.value}>{option.label}</option>
                            ))}
                          </select>
                        </div>
                        <div className="relative" ref={columnsMenuRef}>
                          <button
                            type="button"
                            onClick={() => setColumnsMenuOpen(prev => !prev)}
                            className="inline-flex w-full items-center justify-between gap-2 rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                          >
                            Columns
                            <span aria-hidden="true">▾</span>
                          </button>
                          {columnsMenuOpen ? (
                            <div
                              ref={columnsMenuContentRef}
                              className={`absolute right-0 z-20 w-52 rounded border border-slate-200 bg-white p-2 shadow-lg ${columnsMenuAbove ? 'bottom-full mb-2' : 'top-full mt-2'}`}
                            >
                              <div className="space-y-1">
                                {DOCUMENT_COLUMNS.filter(column => !column.always).map(column => {
                                  const checked = documentColumnsState[column.key] !== false;
                                  return (
                                    <label
                                      key={column.key}
                                      className="flex items-center gap-2 rounded px-2 py-1 text-sm text-slate-600 hover:bg-slate-100"
                                    >
                                      <input
                                        type="checkbox"
                                        className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                                        checked={checked}
                                        onChange={() => handleToggleColumn(column.key)}
                                      />
                                      <span>{column.label}</span>
                                    </label>
                                  );
                                })}
                              </div>
                            </div>
                          ) : null}
                        </div>
                      </div>
                      {selectedCount ? (
                        <div className="text-xs text-slate-500">{selectedCount} selected</div>
                      ) : null}
                    </div>
                    <div className="space-y-3">
                      {documentsError ? (
                        <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700" role="alert">
                          {documentsError}
                        </div>
                      ) : null}
                      {documentsLoading && !documentsError ? (
                        <div className="text-sm text-slate-500">Loading documents…</div>
                      ) : null}
                      {documentsContent}
                    </div>
                  </div>
                </div>
              </section>
            ) : null}

            {workspaceSection === 'templates' ? (
              <section className="rounded-lg border border-slate-200 bg-white p-6">
                <TemplatesManager business={business} onTemplatesUpdated={refreshDocuments} />
              </section>
            ) : null}

            {workspaceSection === 'invoices' ? (
              <section className="rounded-lg border border-slate-200 bg-white p-6">
                <InvoiceLogPanel
                  business={business}
                  onOpenFile={handleOpenDocumentFile}
                  onRevealFile={handleRevealDocument}
                  onDeleteDocument={handleDeleteDocumentRecord}
                />
              </section>
            ) : null}

            {workspaceSection === 'settings' ? (
              <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-700">Business settings</h2>
                  <p className="text-sm text-slate-500">Update folders and review business information.</p>
                </div>

                <div className="rounded border border-slate-200 p-4 flex items-center justify-between gap-3">
                  <div>
                    <h3 className="text-sm font-semibold text-slate-700">Persist UI state</h3>
                    <p className="text-xs text-slate-500">When enabled, restores the exact last view, including selected tabs, open jobsheets and scroll position.</p>
                  </div>
                  <label className="inline-flex items-center gap-2 text-sm text-slate-600">
                    <input type="checkbox" checked={persistUi} onChange={e => setPersistUi(e.target.checked)} />
                    <span>{persistUi ? 'Enabled' : 'Disabled'}</span>
                  </label>
                </div>

                <div className="rounded border border-slate-200 p-4 flex flex-col gap-3">
                  <div>
                    <h3 className="text-sm font-semibold text-slate-700">Documents folder</h3>
                    <p className="text-xs text-slate-500 break-all" title={business.save_path || 'Not configured'}>
                      {business.save_path || 'Not configured'}
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={handleChangeDocumentsFolder}
                      disabled={updatingSavePath}
                      className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60 disabled:cursor-not-allowed"
                    >
                      {updatingSavePath ? 'Updating…' : 'Change folder'}
                    </button>
                    <button
                      type="button"
                      onClick={handleOpenDocumentsFolder}
                      className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                    >
                      Open folder
                    </button>
                  </div>
                </div>

                <div className="rounded border border-slate-200 p-4 flex items-center justify-between gap-3">
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Last invoice number</div>
                    <div className="mt-1 text-sm text-slate-700">{business.last_invoice_number ?? '—'}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={async () => {
                        try {
                          const result = await window.api?.computeFinderInvoiceMax?.({ businessId: business.id });
                          const max = result && Number.isInteger(Number(result.max)) ? Number(result.max) : 0;
                          await window.api?.setLastInvoiceNumber?.(business.id, max);
                          const list = await window.api?.businessSettings?.();
                          if (Array.isArray(list)) {
                            const refreshed = list.find(b => b.id === business.id);
                            if (refreshed && typeof onBusinessUpdate === 'function') {
                              onBusinessUpdate(refreshed);
                            }
                          }
                          setMessage(`Reset last invoice number to ${max} (Finder)`);
                          setTimeout(() => setMessage(''), 2000);
                        } catch (err) {
                          console.error('Failed to reset counter', err);
                          setError(err?.message || 'Unable to reset counter');
                        }
                      }}
                      className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                    >
                      Reset to max (Finder)
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setSetLastDraft(business.last_invoice_number != null ? String(business.last_invoice_number) : '0');
                        setSetLastOpen(true);
                      }}
                      className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                    >
                      Set last…
                    </button>
                  </div>
                </div>

                {setLastOpen ? (
                  <div className="rounded border border-slate-200 p-4 flex flex-col gap-3">
                    <div className="text-sm text-slate-700">Set last invoice number</div>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        min={0}
                        value={setLastDraft}
                        onChange={e => setSetLastDraft(e.target.value)}
                        className="w-32 rounded border border-slate-300 px-2 py-1 text-sm"
                      />
                      <button
                        type="button"
                        onClick={async () => {
                          const val = Number(setLastDraft);
                          if (!Number.isInteger(val) || val < 0) { setError('Enter a non-negative integer'); return; }
                          try {
                            await window.api?.setLastInvoiceNumber?.(business.id, val);
                            const list = await window.api?.businessSettings?.();
                            if (Array.isArray(list)) {
                              const refreshed = list.find(b => b.id === business.id);
                              if (refreshed && typeof onBusinessUpdate === 'function') {
                                onBusinessUpdate(refreshed);
                              }
                            }
                            setSetLastOpen(false);
                            setMessage(`Set last invoice number to ${val}`);
                            setTimeout(() => setMessage(''), 2000);
                          } catch (err) {
                            console.error('Failed to set counter', err);
                            setError(err?.message || 'Unable to set counter');
                          }
                        }}
                        className="inline-flex items-center rounded bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-500"
                      >
                        Save
                      </button>
                      <button
                        type="button"
                        onClick={() => setSetLastOpen(false)}
                        className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                      >
                        Cancel
                      </button>
                    </div>
                  </div>
                ) : null}

                <p className="text-xs text-slate-500">
                  Template management has moved to the “Templates” tab for a simpler workflow. Use it to copy placeholders and replace template files in one place.
                </p>
              </section>
            ) : null}
          </div>
        </div>
      </main>

    </div>
  );
}

// Minimal MCMS workspace: clients + quotes/invoices
function MCMSWorkspace({ business, onBusinessUpdate }) {
  const [clients, setClients] = useState([]);
  const [clientsLoading, setClientsLoading] = useState(true);
  const [clientEditorOpen, setClientEditorOpen] = useState(false);
  const [clientEditorLoading, setClientEditorLoading] = useState(false);
  const [clientEditor, setClientEditor] = useState(null); // { client, emails, phones, addresses }
  const [docs, setDocs] = useState([]);
  const [docsLoading, setDocsLoading] = useState(true);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');

  // Quick-add removed; use full editor modal
  const [creatingDoc, setCreatingDoc] = useState(false);
  const [docType, setDocType] = useState('invoice');
  const [docClientId, setDocClientId] = useState('');
  const [docAmount, setDocAmount] = useState('');
  const [docDueDate, setDocDueDate] = useState('');
  const [items, setItems] = useState([]);
  const [templateHtml, setTemplateHtml] = useState('');
  // Excel template generation state
  const [excelClientId, setExcelClientId] = useState('');
  const [excelAmount, setExcelAmount] = useState('');
  const [excelDueDate, setExcelDueDate] = useState('');
  const [excelTemplatePath, setExcelTemplatePath] = useState('');
  const [excelBusy, setExcelBusy] = useState(false);

  const loadInvoiceDefinition = useCallback(async () => {
    try {
      const defs = await window.api.getDocumentDefinitions(business.id, { includeInactive: true });
      const list = Array.isArray(defs) ? defs : [];
      const def = list.find(d => String(d.key || '').toLowerCase() === 'invoice_balance');
      setExcelTemplatePath(def?.template_path || '');
    } catch (err) { /* ignore */ }
  }, [business.id]);

  useEffect(() => {
    if (workspaceSection === 'invoice') loadInvoiceDefinition();
  }, [workspaceSection, loadInvoiceDefinition]);
  const [workspaceSection, setWorkspaceSection] = useState('dashboard');
  // Client search and Contacts import state
  const [clientSearch, setClientSearch] = useState('');
  const [contactsOpen, setContactsOpen] = useState(false);
  const [contactsLoading, setContactsLoading] = useState(false);
  const [contacts, setContacts] = useState([]);
  const [contactsSearch, setContactsSearch] = useState('');
  const [selectedContacts, setSelectedContacts] = useState(() => new Set());
  const [skipDuplicates, setSkipDuplicates] = useState(true);

  const refreshClients = useCallback(async () => {
    setClientsLoading(true);
    try {
      const list = await window.api.getClients();
      const filtered = Array.isArray(list) ? list.filter(c => !c.business_id || c.business_id === business.id) : [];
      setClients(filtered);
    } catch (err) {
      console.error('Failed to load clients', err);
      setError(err?.message || 'Unable to load clients');
    } finally { setClientsLoading(false); }
  }, [business.id]);

  const refreshDocs = useCallback(async () => {
    setDocsLoading(true);
    try {
      const list = await window.api.getDocuments({ businessId: business.id });
      const filtered = Array.isArray(list) ? list.filter(d => {
        const t = String(d.doc_type || '').toLowerCase();
        return t === 'invoice' || t === 'quote';
      }) : [];
      setDocs(filtered);
    } catch (err) {
      console.error('Failed to load documents', err);
      setError(err?.message || 'Unable to load documents');
    } finally { setDocsLoading(false); }
  }, [business.id]);

  useEffect(() => { refreshClients(); }, [refreshClients]);
  useEffect(() => { refreshDocs(); }, [refreshDocs]);

  const handleAddClient = useCallback((e) => {
    if (e && typeof e.preventDefault === 'function') e.preventDefault();
    setError('');
    setClientEditorOpen(true);
    setClientEditor({
      client: { business_id: business.id, name: '' },
      emails: [{ label: 'Primary', email: '', is_primary: 1 }],
      phones: [{ label: 'Mobile', phone: '', is_primary: 1 }],
      addresses: [{ label: 'Billing', address1: '', address2: '', town: '', postcode: '', country: '', is_primary: 1 }]
    });
  }, [business.id]);

  const openEditClient = useCallback(async (client) => {
    if (!client) return;
    setClientEditorOpen(true);
    setClientEditorLoading(true);
    setError('');
    try {
      const details = await window.api.getClientDetails(client.client_id);
      const base = details?.client || client || {};
      const emails = Array.isArray(details?.emails) && details.emails.length ? details.emails : [{ label: 'Primary', email: base.email || '', is_primary: 1 }];
      const phones = Array.isArray(details?.phones) && details.phones.length ? details.phones : [{ label: 'Mobile', phone: base.phone || '', is_primary: 1 }];
      const addresses = Array.isArray(details?.addresses) && details.addresses.length ? details.addresses : [{ label: 'Billing', address1: base.address1 || '', address2: base.address2 || '', town: base.town || '', postcode: base.postcode || '', country: '' , is_primary: 1 }];
      setClientEditor({ client: base, emails, phones, addresses });
    } catch (err) {
      console.error('Failed to load client details', err);
      setError(err?.message || 'Unable to load client details');
      setClientEditor({ client, emails: [], phones: [], addresses: [] });
    } finally { setClientEditorLoading(false); }
  }, []);

  const closeEditClient = useCallback(() => { setClientEditorOpen(false); setClientEditor(null); }, []);

  const setPrimary = (list, index) => list.map((item, i) => ({ ...item, is_primary: i === index ? 1 : 0 }));

  const saveClientEditor = useCallback(async () => {
    if (!clientEditor || !clientEditor.client) return;
    setError('');
    setClientEditorLoading(true);
    try {
      let id = clientEditor.client.client_id;
      const name = (clientEditor.client.name || '').trim();
      if (!name) throw new Error('Client name is required');
      // Ensure single primary per list in UI
      const normalize = (arr, key) => {
        const any = arr.some((x) => x && (x.is_primary === 1 || x.is_primary === true || x.is_primary === '1'));
        const copy = arr.map(a => ({ ...a }));
        if (!any && copy.length) copy[0].is_primary = 1;
        return copy;
      };
      const emails = normalize(clientEditor.emails || [], 'email');
      const phones = normalize(clientEditor.phones || [], 'phone');
      const addresses = normalize(clientEditor.addresses || [], 'address1');
      if (!Number.isInteger(id)) {
        try {
          id = await window.api.addClient({ business_id: business.id, name });
        } catch (e) {
          throw e;
        }
      }
      await window.api.saveClientDetails(id, { name, emails, phones, addresses });
      setMessage('Client updated');
      setTimeout(() => setMessage(''), 1500);
      closeEditClient();
      refreshClients();
    } catch (err) {
      console.error('Failed to save client', err);
      setError(err?.message || 'Unable to save client');
    } finally { setClientEditorLoading(false); }
  }, [clientEditor, closeEditClient, refreshClients]);

  const deleteClient = useCallback(async (client) => {
    if (!client) return;
    try {
      const ok = window.confirm(`Delete client “${client.name}”? This cannot be undone.`);
      if (!ok) return;
      await window.api.deleteClient(client.client_id);
      setMessage('Client deleted');
      setTimeout(() => setMessage(''), 1200);
      refreshClients();
    } catch (err) {
      console.error('Delete client failed', err);
      setError(err?.message || 'Unable to delete client');
    }
  }, [refreshClients]);

  const handleCreateDocument = useCallback(async () => {
    setCreatingDoc(true);
    setError('');
    try {
      const client = clients.find(c => String(c.client_id) === String(docClientId));
      if (!client) throw new Error('Select a client');
      const computedTotal = Array.isArray(items) && items.length ? items.reduce((sum, it) => {
        const qty = Number(it?.quantity);
        const rate = Number(it?.rate);
        const line = Number.isFinite(Number(it?.amount)) ? Number(it.amount) : (Number.isFinite(qty) && Number.isFinite(rate) ? qty * rate : 0);
        return sum + (Number.isFinite(line) ? line : 0);
      }, 0) : null;
      const amount = (computedTotal != null && Number.isFinite(computedTotal)) ? computedTotal : Number(docAmount);
      if (!Number.isFinite(amount) || amount <= 0) throw new Error('Enter a valid amount or add line items');
      const res = await window.api.createMCMSDocument({
        business_id: business.id,
        doc_type: docType,
        client_override: {
          name: client.name,
          email: client.email,
          phone: client.phone,
          address1: client.address1 || client.address || '',
          address2: client.address2 || '',
          town: client.town || '',
          postcode: client.postcode || ''
        },
        total_amount: amount,
        line_items: items,
        due_date: docDueDate || null
      });
      setMessage(`${docType === 'invoice' ? 'Invoice' : 'Quote'} #${res?.number ?? ''} generated`);
      setTimeout(() => setMessage(''), 1800);
      setDocAmount(''); setDocClientId(''); setDocDueDate(''); setItems([]);
      await refreshDocs();
    } catch (err) {
      console.error('Failed to create document', err);
      setError(err?.message || 'Unable to create document');
    } finally { setCreatingDoc(false); }
  }, [clients, docType, docClientId, docAmount, docDueDate, business.id, refreshDocs, items]);

  const emailDocument = useCallback(async (doc) => {
    try {
      const subject = `${(doc.doc_type || '').toString().toUpperCase()}${doc.number != null ? ` #${doc.number}` : ''} – ${doc.display_client_name || doc.client_name || ''}`.trim();
      const body = '';
      const to = '';
      await window.api.composeMailDraft({ to, subject, body, attachments: [doc.file_path].filter(Boolean) });
    } catch (err) {
      console.error('Compose mail failed', err);
      setError(err?.message || 'Unable to compose email');
    }
  }, []);

  // Local document helpers for the Invoice log in MCMS tabs
  const handleOpenDocumentFile = useCallback(async (filePath) => {
    const path = filePath || '';
    if (!path) { setError('PDF not available'); return; }
    try {
      setError('');
      const res = await window.api?.openPath?.(path);
      if (res && res.ok === false) throw new Error(res.message || 'Unable to open file');
    } catch (err) {
      console.error('Open failed', err);
      setError(err?.message || 'Unable to open file');
    }
  }, []);

  const handleRevealDocument = useCallback(async (filePath) => {
    const path = filePath || '';
    if (!path) { setError('PDF not available'); return; }
    try {
      setError('');
      const res = await window.api?.showItemInFolder?.(path);
      if (res && res.ok === false) throw new Error(res.message || 'Unable to reveal file');
    } catch (err) {
      console.error('Reveal failed', err);
      setError(err?.message || 'Unable to reveal file');
    }
  }, []);

  const handleDeleteDocumentRecord = useCallback(async (doc) => {
    if (!doc || !doc.document_id) return;
    const locked = !!doc.is_locked;
    if (locked) { window.alert('Unlock the record before deleting.'); return; }
    const removeFile = window.confirm('Also delete the PDF file from disk?');
    try {
      await window.api?.deleteDocument?.(doc.document_id, { removeFile });
      setMessage('Deleted');
      setTimeout(() => setMessage(''), 1200);
      refreshDocs();
    } catch (err) {
      console.error('Delete failed', err);
      setError(err?.message || 'Unable to delete');
    }
  }, [refreshDocs]);

  const documentsContent = (
    <div className="space-y-2">
      {docsLoading ? <div className="text-sm text-slate-500">Loading…</div> : null}
      {docs.map(row => {
        const number = row.number != null ? `#${row.number}` : '';
        const label = `${(row.doc_type || '').toString().toUpperCase()} ${number}`.trim();
        return (
          <div key={row.document_id} className="flex items-center justify-between rounded border border-slate-200 px-3 py-2">
            <div className="flex flex-col">
              <div className="text-sm font-medium text-slate-700">{label} — {row.display_client_name || row.client_name || ''}</div>
              <div className="text-xs text-slate-500">{row.file_path || 'No file yet'}</div>
            </div>
            <div className="flex items-center gap-2">
              {row.file_path ? (
                <>
                  <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={() => window.api.openPath(row.file_path)}>Open</button>
                  <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={() => window.api.showItemInFolder(row.file_path)}>Reveal</button>
                  <button className="text-xs px-2 py-1 border rounded border-indigo-200 text-indigo-600 hover:bg-indigo-50" onClick={() => emailDocument(row)}>Email</button>
                </>
              ) : null}
            </div>
          </div>
        );
      })}
    </div>
  );

  

  const editor = clientEditor || { client: {}, emails: [], phones: [], addresses: [] };
  const onChangeList = (key, updater) => setClientEditor(prev => prev ? { ...prev, [key]: updater(Array.isArray(prev[key]) ? prev[key] : []) } : prev);
  const renderList = (title, key, columns) => {
    const list = Array.isArray(editor[key]) ? editor[key] : [];
    return (
      <div className="space-y-2">
        <div className="text-sm font-semibold text-slate-700">{title}</div>
        {list.map((row, idx) => (
          <div key={`${key}-${idx}`} className="flex flex-wrap items-end gap-2">
            {columns.map(col => (
              <div key={col.key} className="flex flex-col">
                <label className="text-[11px] text-slate-500">{col.label}</label>
                <input value={row[col.key] || ''} onChange={e => onChangeList(key, (arr) => { const next = [...arr]; next[idx] = { ...next[idx], [col.key]: e.target.value }; return next; })} className="border rounded px-2 py-1 text-sm" />
              </div>
            ))}
            <label className="text-xs text-slate-600 inline-flex items-center gap-1">
              <input type="checkbox" checked={!!row.is_primary} onChange={() => onChangeList(key, (arr) => setPrimary(arr, idx))} /> Primary
            </label>
            <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={() => onChangeList(key, (arr) => arr.filter((_, i) => i !== idx))}>Remove</button>
          </div>
        ))}
        <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={() => onChangeList(key, (arr) => arr.concat([columns.reduce((o, c) => ({ ...o, [c.key]: '' }), { is_primary: list.length === 0 ? 1 : 0 })]))}>Add</button>
      </div>
    );
  };

  const editorModal = clientEditorOpen ? (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div className="bg-white rounded shadow-xl w-full max-w-3xl p-6 space-y-4">
        <div className="flex items-center justify-between">
          <div>
            <div className="text-lg font-semibold text-slate-800">Edit client</div>
            <div className="text-xs text-slate-500">Update contact methods and addresses</div>
          </div>
          <div className="flex items-center gap-2">
            <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={async ()=>{
              setContactsOpen(true);
              if (contacts.length) return;
              setContactsLoading(true);
              try {
                const res = await window.api.listAppleContacts();
                const list = (res && res.ok && Array.isArray(res.contacts)) ? res.contacts : [];
                setContacts(list);
              } catch (err) {
                console.error('Failed to list Apple Contacts', err);
                setError(err?.message || 'Unable to read Apple Contacts');
              } finally { setContactsLoading(false); }
            }}>Import from Contacts</button>
            <button className="text-sm px-3 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={closeEditClient}>Close</button>
          </div>
        </div>
        <div className="space-y-3">
          <div className="flex flex-col">
            <label className="text-[11px] text-slate-500">Name</label>
            <input value={editor.client?.name || ''} onChange={e => setClientEditor(prev => prev ? { ...prev, client: { ...(prev.client || {}), name: e.target.value } } : prev)} className="border rounded px-2 py-1 text-sm" />
          </div>
          {renderList('Emails', 'emails', [
            { key: 'label', label: 'Label' },
            { key: 'email', label: 'Email' }
          ])}
          {renderList('Phones', 'phones', [
            { key: 'label', label: 'Label' },
            { key: 'phone', label: 'Phone' }
          ])}
          {renderList('Addresses', 'addresses', [
            { key: 'label', label: 'Label' },
            { key: 'address1', label: 'Address line 1' },
            { key: 'address2', label: 'Address line 2' },
            { key: 'town', label: 'Town/City' },
            { key: 'postcode', label: 'Postcode' },
            { key: 'country', label: 'Country' }
          ])}
        </div>
        <div className="flex items-center justify-end gap-2">
          <button className="text-sm px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={closeEditClient}>Cancel</button>
          <button className="text-sm px-3 py-2 rounded bg-indigo-600 text-white hover:bg-indigo-500 disabled:opacity-50" disabled={clientEditorLoading} onClick={saveClientEditor}>{clientEditorLoading ? 'Saving…' : 'Save changes'}</button>
        </div>
      </div>
    </div>
  ) : null;

  // Contacts import modal and helpers
  const contactsList = Array.isArray(contacts) ? contacts : [];
  const filteredContacts = contactsList.filter(c => {
    if (!contactsSearch.trim()) return true;
    const q = contactsSearch.trim().toLowerCase();
    const hay = [c.name, (c.emails||[]).map(e=>e.value).join(' '), (c.phones||[]).map(p=>p.value).join(' ')].join(' ').toLowerCase();
    return hay.includes(q);
  });
  const toggleContact = (id) => setSelectedContacts(prev => { const next = new Set(Array.from(prev)); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  const setAllContacts = (all) => setSelectedContacts(all ? new Set(filteredContacts.map(c=>c.id||c.name)) : new Set());
  const importSelectedContacts = async () => {
    setError('');
    try {
      const ids = Array.from(selectedContacts);
      if (!ids.length) { setContactsOpen(false); return; }
      const list = contactsList.filter(c => ids.includes(c.id || c.name));
      for (const c of list) {
        const name = (c.name || '').trim() || ((c.firstName||'') + ' ' + (c.lastName||'')).trim();
        if (!name) continue;
        let existing = null;
        try { existing = await window.api.getClientByName(business.id, name); } catch (_) {}
        if (existing && skipDuplicates) continue;
        let clientId = existing ? existing.client_id : null;
        if (!clientId) {
          try { clientId = await window.api.addClient({ business_id: business.id, name }); } catch (e) { clientId = null; }
        }
        if (!clientId) continue;
        const emails = Array.isArray(c.emails) ? c.emails.filter(e => e && e.value).map((e,i) => ({ label: e.label || '', email: e.value || '', is_primary: i === 0 ? 1 : 0 })) : [];
        const phones = Array.isArray(c.phones) ? c.phones.filter(p => p && p.value).map((p,i) => ({ label: p.label || '', phone: p.value || '', is_primary: i === 0 ? 1 : 0 })) : [];
        const addresses = Array.isArray(c.addresses) ? c.addresses.map((a,i) => ({ label: a.label || '', address1: a.street || '', address2: '', town: a.city || '', postcode: a.zip || '', country: a.country || '', is_primary: i === 0 ? 1 : 0 })) : [];
        try { await window.api.saveClientDetails(clientId, { name, emails, phones, addresses }); } catch (_) {}
      }
      setContactsOpen(false);
      setSelectedContacts(new Set());
      setMessage('Import complete');
      setTimeout(() => setMessage(''), 1500);
      refreshClients();
    } catch (err) {
      console.error('Import failed', err);
      setError(err?.message || 'Unable to import contacts');
    }
  };

  const importModal = contactsOpen ? (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div className="bg-white rounded shadow-xl w-full max-w-4xl p-6 space-y-4">
        <div className="flex items-center justify-between">
          <div>
            <div className="text-lg font-semibold text-slate-800">Import from Apple Contacts</div>
            <div className="text-xs text-slate-500">Select which contacts to import</div>
          </div>
          <button className="text-sm px-3 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setContactsOpen(false)}>Close</button>
        </div>
        <div className="flex items-center justify-between">
          <input value={contactsSearch} onChange={e=>setContactsSearch(e.target.value)} placeholder="Search contacts…" className="text-sm border rounded px-2 py-1" />
          <div className="flex items-center gap-2">
            <label className="text-xs text-slate-600 inline-flex items-center gap-1"><input type="checkbox" checked={skipDuplicates} onChange={e=>setSkipDuplicates(e.target.checked)} /> Skip duplicates</label>
            <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setAllContacts(true)}>Select all</button>
            <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setAllContacts(false)}>Clear</button>
          </div>
        </div>
        <div className="border rounded max-h-[50vh] overflow-auto">
          {contactsLoading ? (
            <div className="p-3 text-sm text-slate-500">Loading contacts…</div>
          ) : filteredContacts.length === 0 ? (
            <div className="p-3 text-sm text-slate-500">No contacts found. macOS may require Contacts permission for Automation. You can retry or import a vCard (.vcf) file.</div>
          ) : (
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-700">
                  <th className="px-3 py-2 w-10"></th>
                  <th className="px-3 py-2 text-left">Name</th>
                  <th className="px-3 py-2 text-left">Emails</th>
                  <th className="px-3 py-2 text-left">Phones</th>
                </tr>
              </thead>
              <tbody>
                {filteredContacts.map(c => {
                  const id = c.id || c.name;
                  const checked = selectedContacts.has(id);
                  return (
                    <tr key={id} className="border-t">
                      <td className="px-3 py-2"><input type="checkbox" checked={checked} onChange={()=>toggleContact(id)} /></td>
                      <td className="px-3 py-2">{c.name || `${c.firstName||''} ${c.lastName||''}`}</td>
                      <td className="px-3 py-2">{(c.emails||[]).map(e=>e.value).filter(Boolean).join(', ')}</td>
                      <td className="px-3 py-2">{(c.phones||[]).map(p=>p.value).filter(Boolean).join(', ')}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
        <div className="flex items-center justify-between">
          <div>
            <button
              className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50"
              onClick={async ()=>{
                try {
                  const file = await window.api.chooseFile({ title: 'Select vCard file', filters: [{ name: 'vCard', extensions: ['vcf', 'vcard'] }] });
                  if (!file) return;
                  // Simple vCard parse in main process is not available; instead, quick-read via preload is not present.
                  // Fallback UX: suggest dragging contacts into Contacts app selection then retry.
                  setError('vCard import will be added next. For now, select contacts in Contacts and retry.');
                } catch (err) {
                  setError(err?.message || 'Unable to import vCard');
                }
              }}
            >Import vCard…</button>
          </div>
          <div className="flex items-center gap-2">
            <button className="text-sm px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setContactsOpen(false)}>Cancel</button>
            <button className="text-sm px-3 py-2 rounded bg-indigo-600 text-white hover:bg-indigo-500 disabled:opacity-50" disabled={contactsLoading || selectedContacts.size === 0} onClick={importSelectedContacts}>Import selected</button>
          </div>
        </div>
      </div>
    </div>
  ) : null;

  // Render MCMS layout + editor modal
  return (
    <div className="min-h-screen bg-slate-100">
      <ToastOverlay notices={[
        error ? { id: 'mcms-error', tone: 'error', text: error } : null,
        message ? { id: 'mcms-message', tone: 'success', text: message } : null
      ].filter(Boolean)} />
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold text-slate-800">{business.business_name}</h1>
            <p className="text-sm text-slate-500">Clients, quotes and invoices.</p>
          </div>
          {/* Switch business removed */}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        <nav className="flex gap-2">
          {['clients','invoice','settings'].map(key => (
            <button key={key} onClick={() => setWorkspaceSection(key)} className={`text-sm px-3 py-2 rounded border ${workspaceSection===key? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>{key.charAt(0).toUpperCase()+key.slice(1)}</button>
          ))}
        </nav>

        {workspaceSection === 'clients' ? (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-slate-700">Clients</h2>
                <div className="flex items-center gap-2">
                  <input value={clientSearch} onChange={e=>setClientSearch(e.target.value)} placeholder="Search…" className="text-xs border rounded px-2 py-1" />
                  <button onClick={refreshClients} className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50">Refresh</button>
                  <button onClick={async ()=>{
                    setContactsOpen(true);
                    if (contacts.length) return;
                    setContactsLoading(true);
                    try {
                      const res = await window.api.listAppleContacts();
                      const list = (res && res.ok && Array.isArray(res.contacts)) ? res.contacts : [];
                      setContacts(list);
                    } catch (err) {
                      console.error('Failed to list Apple Contacts', err);
                      setError(err?.message || 'Unable to read Apple Contacts');
                    } finally { setContactsLoading(false); }
                  }} className="hidden text-xs px-2 py-1 border rounded border-indigo-200 text-indigo-600 hover:bg-indigo-50">Import from Contacts</button>
                  <button onClick={handleAddClient} className="text-xs px-2 py-1 border rounded border-indigo-200 text-indigo-600 hover:bg-indigo-50">Add contact</button>
                </div>
              </div>
              
              <div className="divide-y divide-slate-100">
                {clientsLoading ? <div className="text-sm text-slate-500">Loading…</div> : null}
                {clients
                  .filter(c => {
                    if (!clientSearch.trim()) return true;
                    const q = clientSearch.trim().toLowerCase();
                    const hay = [c.name, c.email, c.phone, c.address, c.town, c.postcode].filter(Boolean).join(' ').toLowerCase();
                    return hay.includes(q);
                  })
                  .map(c => (
                  <div key={c.client_id} className="py-2 text-sm flex items-center justify-between">
                    <div>
                      <div className="font-medium text-slate-700">{c.name}</div>
                      <div className="text-xs text-slate-500">{[c.email, c.phone].filter(Boolean).join(' · ')}</div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={() => openEditClient(c)}>Edit</button>
                      <button className="text-xs px-2 py-1 border rounded border-red-200 text-red-600 hover:bg-red-50" onClick={() => deleteClient(c)}>Delete</button>
                    </div>
                  </div>
                ))}
              </div>
            </section>
          </div>
        ) : null}

        {workspaceSection === 'invoice' ? (
          <div className="space-y-6">
            <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-slate-700">Generate from Excel template</h2>
                <div className="text-xs text-slate-500">{excelTemplatePath ? `Template: ${excelTemplatePath}` : 'No template set'}</div>
              </div>
              <details className="rounded border border-slate-200">
                <summary className="cursor-pointer px-3 py-2 text-sm font-medium text-slate-700 bg-slate-50">Edit Excel template (beta)</summary>
                <div className="p-3">
                  <ExcelTemplateEditor initialPath={excelTemplatePath} onSaved={()=>{ setMessage('Template saved'); setTimeout(()=>setMessage(''), 1200); }} />
                </div>
              </details>
              <div className="flex flex-wrap items-end gap-2">
                <div className="flex flex-col min-w-[220px]">
                  <label className="text-xs text-slate-500">Client</label>
                  <select value={excelClientId} onChange={e=>setExcelClientId(e.target.value)} className="border rounded px-2 py-1 text-sm">
                    <option value="">Select…</option>
                    {clients.map(c => (<option key={c.client_id} value={c.client_id}>{c.name}</option>))}
                  </select>
                </div>
                <div className="flex flex-col">
                  <label className="text-xs text-slate-500">Amount</label>
                  <input type="number" step="0.01" value={excelAmount} onChange={e=>setExcelAmount(e.target.value)} className="border rounded px-2 py-1 text-sm" placeholder="0.00" />
                </div>
                <div className="flex flex-col">
                  <label className="text-xs text-slate-500">Due date</label>
                  <input type="date" value={excelDueDate} onChange={e=>setExcelDueDate(e.target.value)} className="border rounded px-2 py-1 text-sm" />
                </div>
                <button className="text-xs px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={async ()=>{
                  try {
                    const file = await window.api.chooseFile({ title: 'Select invoice template (xlsx)', filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }] });
                    if (!file) return;
                    await window.api.saveDocumentDefinition(business.id, { key: 'invoice_balance', doc_type: 'invoice', label: 'Invoice – Balance', template_path: file, is_active: 1, is_locked: 0 });
                    setExcelTemplatePath(file);
                    setMessage('Template set'); setTimeout(()=>setMessage(''), 1200);
                  } catch (err) { setError(err?.message || 'Unable to set template'); }
                }}>Set template…</button>
                <button disabled={excelBusy} className="text-xs px-3 py-2 rounded bg-indigo-600 text-white hover:bg-indigo-500 disabled:opacity-50" onClick={async ()=>{
                  setError(''); setExcelBusy(true);
                  try {
                    const client = clients.find(c => String(c.client_id) === String(excelClientId));
                    if (!client) throw new Error('Select a client');
                    const amt = Number(excelAmount);
                    if (!Number.isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
                    let tpl = excelTemplatePath;
                    if (!tpl) {
                      const defs = await window.api.getDocumentDefinitions(business.id, { includeInactive: true });
                      const def = (Array.isArray(defs)?defs:[]).find(d => String(d.key||'').toLowerCase()==='invoice_balance');
                      tpl = def?.template_path || '';
                    }
                    if (!tpl) throw new Error('Please set an Excel template first');
                    const res = await window.api.createNumberedDocument({
                      business_id: business.id,
                      doc_type: 'invoice',
                      definition_key: 'invoice_balance',
                      client_override: {
                        name: client.name, email: client.email, phone: client.phone,
                        address1: client.address1 || client.address || '', address2: client.address2 || '', town: client.town || '', postcode: client.postcode || ''
                      },
                      total_amount: amt,
                      due_date: excelDueDate || null
                    });
                    if (!res || !res.file_path) throw new Error('PDF not created');
                    setMessage(`Invoice #${res?.number ?? ''} generated`); setTimeout(()=>setMessage(''), 1500);
                    setExcelAmount(''); setExcelClientId(''); setExcelDueDate('');
                    refreshDocs();
                  } catch (err) { console.error(err); setError(err?.message || 'Unable to generate invoice'); }
                  finally { setExcelBusy(false); }
                }}>{excelBusy ? 'Generating…' : 'Generate Invoice'}</button>
              </div>
            </section>
            
            <section className="rounded-lg border border-slate-200 bg-white p-6">
              <InvoiceLogPanel
                business={business}
                onOpenFile={handleOpenDocumentFile}
                onRevealFile={handleRevealDocument}
                onDeleteDocument={handleDeleteDocumentRecord}
              />
            </section>
          </div>
        ) : null}

        {workspaceSection === 'quote' ? (
          <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
            <div className="flex items-center justify-between">
              <h2 className="text-lg font-semibold text-slate-700">New Quote</h2>
            </div>
            <div className="flex flex-wrap items-end gap-2">
              <div className="flex flex-col min-w-[200px]">
                <label className="text-xs text-slate-500">Client</label>
                <select value={docClientId} onChange={e=>setDocClientId(e.target.value)} className="border rounded px-2 py-1 text-sm">
                  <option value="">Select…</option>
                  {clients.map(c => (
                    <option key={c.client_id} value={c.client_id}>{c.name}</option>
                  ))}
                </select>
              </div>
              {items.length === 0 ? (
                <div className="flex flex-col">
                  <label className="text-xs text-slate-500">Amount</label>
                  <input type="number" step="0.01" value={docAmount} onChange={e=>setDocAmount(e.target.value)} className="border rounded px-2 py-1 text-sm" placeholder="0.00" />
                </div>
              ) : (
                <div className="flex flex-col">
                  <label className="text-xs text-slate-500">Total</label>
                  <div className="px-2 py-1 text-sm">£{(items.reduce((s,it)=>{const q=Number(it.quantity);const r=Number(it.rate);const line=Number.isFinite(Number(it.amount))?Number(it.amount):(Number.isFinite(q)&&Number.isFinite(r)?q*r:0);return s+(Number.isFinite(line)?line:0);},0)).toFixed(2)}</div>
                </div>
              )}
              <div className="flex flex-col">
                <label className="text-xs text-slate-500">Valid until</label>
                <input type="date" value={docDueDate} onChange={e=>setDocDueDate(e.target.value)} className="border rounded px-2 py-1 text-sm" />
              </div>
              <button disabled={creatingDoc} onClick={async ()=>{
                setCreatingDoc(true);
                setError('');
                try{
                  const client = clients.find(c => String(c.client_id) === String(docClientId));
                  if (!client) throw new Error('Select a client');
                  const computed = items.length ? items.reduce((s,it)=>{const qa=Number(it.quantity), ra=Number(it.rate); const ln = Number.isFinite(Number(it.amount))?Number(it.amount):(Number.isFinite(qa)&&Number.isFinite(ra)?qa*ra:0); return s+(Number.isFinite(ln)?ln:0);},0) : null;
                  const amount = (computed != null && Number.isFinite(computed)) ? computed : Number(docAmount);
                  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Enter a valid amount or add line items');
                  const res = await window.api.createMCMSDocument({
                    business_id: business.id,
                    doc_type: 'quote',
                    client_override: {
                      name: client.name, email: client.email, phone: client.phone,
                      address1: client.address1 || client.address || '', address2: client.address2 || '', town: client.town || '', postcode: client.postcode || ''
                    },
                    total_amount: amount,
                    line_items: items,
                    due_date: docDueDate || null
                  });
                  setMessage(`Quote #${res?.number ?? ''} generated`);
                  setTimeout(()=>setMessage(''), 1800);
                  setDocClientId(''); setDocAmount(''); setDocDueDate(''); setItems([]);
                  await refreshDocs();
                }catch(err){ console.error(err); setError(err?.message || 'Unable to generate quote'); }
                finally{ setCreatingDoc(false); }
              }} className="text-xs px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-50">{creatingDoc ? 'Generating…' : 'Generate Quote'}</button>
            </div>

            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div className="text-sm font-semibold text-slate-700">Line items</div>
                <div className="flex gap-2">
                  <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setItems(arr=>arr.concat([{ item_type:'gig', description:'Performance fee', quantity:1, unit:'each', rate:0 }]))}>Add gig fee</button>
                  <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setItems(arr=>arr.concat([{ item_type:'studio', description:'Studio time', quantity:1, unit:'hours', rate:0 }]))}>Add studio time</button>
                  <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>setItems(arr=>arr.concat([{ item_type:'expense', description:'Expense', quantity:1, unit:'item', rate:0 }]))}>Add expense</button>
                </div>
              </div>
              {items.length ? (
                <div className="border rounded overflow-hidden">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-slate-50 text-slate-700">
                        <th className="px-2 py-1 text-left">Type</th>
                        <th className="px-2 py-1 text-left">Description</th>
                        <th className="px-2 py-1 text-right" style={{width:'90px'}}>Qty/Hrs</th>
                        <th className="px-2 py-1 text-left" style={{width:'80px'}}>Unit</th>
                        <th className="px-2 py-1 text-right" style={{width:'120px'}}>Rate</th>
                        <th className="px-2 py-1 text-right" style={{width:'120px'}}>Line total</th>
                        <th className="px-2 py-1"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {items.map((it, idx) => {
                        const qty = Number(it.quantity);
                        const rate = Number(it.rate);
                        const line = Number.isFinite(Number(it.amount)) ? Number(it.amount) : (Number.isFinite(qty) && Number.isFinite(rate) ? qty * rate : 0);
                        return (
                          <tr key={`qt-${idx}`} className="border-t">
                            <td className="px-2 py-1">
                              <select value={it.item_type||''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = { ...next[idx], item_type: e.target.value }; return next; })} className="border rounded px-1 py-0.5 text-xs">
                                <option value="gig">Gig</option>
                                <option value="studio">Studio</option>
                                <option value="expense">Expense</option>
                                <option value="custom">Custom</option>
                              </select>
                            </td>
                            <td className="px-2 py-1"><input value={it.description||''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = { ...next[idx], description: e.target.value }; return next; })} className="w-full border rounded px-2 py-1 text-xs" placeholder="Description" /></td>
                            <td className="px-2 py-1 text-right"><input type="number" step="0.01" value={Number.isFinite(qty)?qty:''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = { ...next[idx], quantity: e.target.value }; return next; })} className="w-20 border rounded px-2 py-1 text-xs text-right" /></td>
                            <td className="px-2 py-1"><input value={it.unit||''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = { ...next[idx], unit: e.target.value }; return next; })} className="w-20 border rounded px-2 py-1 text-xs" /></td>
                            <td className="px-2 py-1 text-right"><input type="number" step="0.01" value={Number.isFinite(rate)?rate:''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = { ...next[idx], rate: e.target.value }; return next; })} className="w-28 border rounded px-2 py-1 text-xs text-right" /></td>
                            <td className="px-2 py-1 text-right">£{Number.isFinite(line)?line.toFixed(2):'0.00'}</td>
                            <td className="px-2 py-1 text-right"><button className="text-xs px-2 py-1 border rounded border-red-200 text-red-600 hover:bg-red-50" onClick={()=>setItems(arr=>arr.filter((_,i)=>i!==idx))}>Remove</button></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ) : null}
            </div>

            <div className="space-y-3">
              <div className="text-sm font-semibold text-slate-700">Recent quotes</div>
              <div className="divide-y divide-slate-100">
                {docs
                  .filter(d => String(d.doc_type || '').toLowerCase() === 'quote')
                  .map(d => (
                    <div key={`q-${d.document_id}`} className="py-2 text-sm flex items-center justify-between">
                      <div>
                        <div className="font-medium text-slate-700">Quote #{d.number != null ? d.number : ''} — {d.display_client_name || d.client_name || ''}</div>
                        <div className="text-xs text-slate-500 break-all">{d.file_path || 'No file yet'}</div>
                      </div>
                      <div className="flex items-center gap-2">
                        <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>window.api.openPath(d.file_path)}>Open</button>
                        <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>window.api.showItemInFolder(d.file_path)}>Reveal</button>
                        <button className="text-xs px-2 py-1 border rounded border-red-200 text-red-600 hover:bg-red-50" onClick={async ()=>{ try { await window.api.deleteDocument(d.document_id, { removeFile: true }); setMessage('Quote deleted'); setTimeout(()=>setMessage(''), 1200); refreshDocs(); } catch(err){ setError(err?.message || 'Unable to delete'); } }}>Delete</button>
                      </div>
                    </div>
                  ))}
              </div>
            </div>
          </div>
        ) : null}

        {workspaceSection === 'invoices' ? (
          <section className="rounded-lg border border-slate-200 bg-white p-6">
            <InvoiceLogPanel
              business={business}
              onOpenFile={async (filePath) => {
                if (!filePath) return;
                try { const res = await window.api.openPath(filePath); if (res && res.ok === false) { throw new Error(res.message || 'Unable to open'); } } catch (err) { setError(err?.message || 'Unable to open file'); }
              }}
              onRevealFile={async (filePath) => {
                if (!filePath) return;
                try { const res = await window.api.showItemInFolder(filePath); if (res && res.ok === false) { throw new Error(res.message || 'Unable to reveal'); } } catch (err) { setError(err?.message || 'Unable to reveal file'); }
              }}
              onDeleteDocument={async (doc) => {
                if (!doc || !doc.document_id) return;
                const locked = !!doc.is_locked;
                if (locked) { window.alert('Unlock the record before deleting.'); return; }
                const removeFile = window.confirm('Also delete the PDF file from disk?');
                try { await window.api.deleteDocument(doc.document_id, { removeFile }); setMessage('Invoice deleted'); setTimeout(()=>setMessage(''), 1200); } catch (err) { setError(err?.message || 'Unable to delete'); }
              }}
            />
          </section>
        ) : null}

        {workspaceSection === 'settings' ? (
          <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
            <div>
              <h2 className="text-lg font-semibold text-slate-700">Business settings</h2>
              <p className="text-sm text-slate-500">Update folders and review business information.</p>
            </div>
            <div className="rounded border border-slate-200 p-4 flex items-center justify-between gap-3">
              <div>
                <h3 className="text-sm font-semibold text-slate-700">Documents folder</h3>
                <p className="text-xs text-slate-500 break-all">{business.save_path || 'Not configured'}</p>
              </div>
            </div>
          </section>
        ) : null}
      </main>

      {editorModal}
      {importModal}
    </div>
  );
}

function JobsheetDocumentsPanel({
  jobsheetId,
  documents,
  documentDefinitions,
  loading,
  definitionsLoading,
  error,
  onRefresh,
  onGenerate,
  onRegenerate,
  generatingKey,
  workbookDefinition,
  onOpenTemplate,
  onEditTemplate,
  onOpenOutputFolder,
  onOpenOutputFile,
  onOpenFile,
  onRevealFile,
  onDelete,
  onExportPdf,
  documentFolder,
  lastOutputPath
}) {
  if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) return null;

  const list = Array.isArray(documents) ? documents : [];
  const workbooks = list.filter(doc => (doc?.doc_type || '').toLowerCase() === 'workbook');
  const pdfExports = list.filter(doc => (doc?.doc_type || '').toLowerCase().includes('pdf'));
  const otherDocuments = list.filter(doc => !workbooks.includes(doc) && !pdfExports.includes(doc));
  // minimal generation actions
  const workbookDefs = useMemo(() => (
    Array.isArray(documentDefinitions)
      ? documentDefinitions.filter(def => (def.doc_type || '').toLowerCase() === 'workbook')
      : []
  ), [documentDefinitions]);
  const lastWorkbook = workbooks.length ? workbooks[0] : null;
  const canExportPdfs = Boolean(lastWorkbook && lastWorkbook.file_path);
  const [generatingAll, setGeneratingAll] = useState(false);
  const handleGenerateAll = useCallback(async () => {
    if (!onGenerate || !jobsheetId) return;
    const ready = workbookDefs.filter(def => def && def.template_path);
    if (!ready.length) return;
    try {
      setGeneratingAll(true);
      for (const def of ready) {
        // eslint-disable-next-line no-await-in-loop
        await onGenerate(def.key);
      }
    } finally {
      setGeneratingAll(false);
    }
  }, [onGenerate, jobsheetId, workbookDefs]);
  const documentLabel = (doc) => {
    return doc?.display_label
      || doc?.definition_label
      || (doc?.definition_key ? startCaseKey(doc.definition_key) : 'Document');
  };

  // generation UI removed; keep outputs only

  const renderDocumentRow = (doc, options = {}) => {
    if (!doc) return null;
    const { showExport = false, showRegenerate = false } = options;
    const rowKey = doc?.document_id != null
      ? `doc-${doc.document_id}`
      : doc?.file_path
        ? `path-${doc.file_path}`
        : `${doc?.doc_type || 'doc'}-${documentLabel(doc)}`;
    const title = documentLabel(doc);
    const fileName = doc?.file_name
      || (doc?.file_path ? doc.file_path.split(/[\\/]+/).filter(Boolean).pop() : '');
    const createdDisplay = doc?.created_at ? formatTimestampDisplay(doc.created_at) : '—';
    const tooltip = doc?.file_path || title;
    const missingFile = doc?.file_available === false;
    const disableFileActions = !doc?.file_path || missingFile;
    const canExport = showExport && typeof onExportPdf === 'function';
    const canRegenerate = showRegenerate && typeof onRegenerate === 'function' && doc?.definition_key;
    const isRegenerating = canRegenerate && generatingKey === doc.definition_key;

    return (
      <tr key={rowKey}>
        <td className="px-3 py-2 align-top">
          <div className="flex items-start gap-3">
            <span className="text-lg" role="img" aria-label={title}>{getDocumentIcon(doc.doc_type)}</span>
            <div className="space-y-1">
              <div
                className="text-xs font-medium text-slate-700 truncate"
                style={{ maxWidth: '24rem' }}
                title={tooltip}
              >
                {fileName || title}
              </div>
              {title && title !== fileName ? (
                <div className="text-[11px] text-slate-500">{title}</div>
              ) : null}
            {null}
              {missingFile ? (
                <div className="text-[11px] text-rose-600">File missing on disk</div>
              ) : null}
            </div>
          </div>
        </td>
        <td className="px-3 py-2 align-top text-sm text-slate-600">{createdDisplay}</td>
        <td className="px-3 py-2 align-top">
          <div className="flex flex-wrap justify-end gap-2">
            <button
              type="button"
              onClick={() => onOpenFile?.(doc.file_path)}
              disabled={disableFileActions}
              className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              Open
            </button>
            <button
              type="button"
              onClick={() => onRevealFile?.(doc.file_path)}
              disabled={disableFileActions}
              className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              Reveal in Finder
            </button>
            {canExport ? (
              <button
                type="button"
                onClick={() => onExportPdf?.(doc)}
                disabled={disableFileActions}
                className="inline-flex items-center rounded border border-indigo-200 px-2 py-1 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-60"
              >
                Export PDF
              </button>
            ) : null}
            {canRegenerate ? (
              <button
                type="button"
                onClick={() => onRegenerate?.(doc.definition_key, doc)}
                disabled={isRegenerating}
                className="inline-flex items-center rounded border border-indigo-200 px-2 py-1 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-60"
              >
                {isRegenerating ? 'Regenerating…' : 'Regenerate'}
              </button>
            ) : null}
            {onDelete ? (
              <button
                type="button"
                onClick={() => onDelete?.(doc)}
                disabled={doc?.document_id == null}
                className="inline-flex items-center rounded border border-rose-200 px-2 py-1 text-xs font-medium text-rose-600 hover:bg-rose-50 disabled:cursor-not-allowed disabled:opacity-60"
              >
                Delete
              </button>
            ) : null}
          </div>
        </td>
      </tr>
    );
  };

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
      <div className="mb-2 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div className="min-w-0">
          <h2 className="text-lg font-semibold text-slate-700">Documents</h2>
          {documentFolder ? (
            <div className="text-xs text-slate-500 break-all">{documentFolder}</div>
          ) : null}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex flex-wrap items-center gap-2">
            {workbookDefs.map(def => {
              const disabled = !jobsheetId || !def.template_path || definitionsLoading;
              return (
                <button
                  key={def.key}
                  type="button"
                  onClick={() => onGenerate?.(def.key)}
                  disabled={disabled}
                  className="inline-flex items-center rounded border border-slate-300 px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                  title={def.template_path || 'No template configured'}
                >
                  Generate {def.label || startCaseKey(def.key)}
                </button>
              );
            })}
          </div>
          <button
            type="button"
            onClick={handleGenerateAll}
            disabled={generatingAll || definitionsLoading || !jobsheetId || workbookDefs.every(d => !d.template_path)}
            className="inline-flex items-center rounded border border-indigo-200 px-3 py-1.5 text-xs font-semibold text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {generatingAll ? 'Generating…' : 'Generate all'}
          </button>
          <button
            type="button"
            onClick={() => onExportPdf?.(lastWorkbook)}
            disabled={!canExportPdfs}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
          >
            Export PDFs
          </button>
        </div>
      </div>

      {error ? (
        <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-600">{error}</div>
      ) : null}

      {/* Document generation UI removed as requested */}

      <div className="space-y-4">

        {workbooks.length === 0 && pdfExports.length === 0 && otherDocuments.length === 0 && !loading ? (
          <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-500">
            No documents generated yet.
          </div>
        ) : null}

        {workbooks.length > 0 ? (
          <div className="space-y-2">
            <div className="text-sm font-medium text-slate-600">Workbooks</div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-100 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">Workbook</th>
                    <th className="px-3 py-2 text-left">Created</th>
                    <th className="px-3 py-2 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {workbooks.map(doc => renderDocumentRow(doc, { showExport: true, showRegenerate: true }))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null}

        {pdfExports.length > 0 ? (
          <div className="space-y-2">
            <div className="text-sm font-medium text-slate-600">Exports</div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-100 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">PDF</th>
                    <th className="px-3 py-2 text-left">Created</th>
                    <th className="px-3 py-2 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {pdfExports.map(doc => renderDocumentRow(doc))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null}

        {otherDocuments.length > 0 ? (
          <div className="space-y-2">
            <div className="text-sm font-medium text-slate-600">Other documents</div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-100 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">Document</th>
                    <th className="px-3 py-2 text-left">Created</th>
                    <th className="px-3 py-2 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {otherDocuments.map(doc => renderDocumentRow(doc))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null}
      </div>
    </section>
  );
}

function JobsheetEditorWindow({
  businessId,
  businessName,
  initialJobsheetId,
  variant = 'window',
  targetJobsheetId,
  onRequestClose
}) {
  const isInline = variant === 'inline';
  const resolveJobsheetId = (value) => {
    if (value === undefined || value === null) return null;
    if (value === '' || value === 'new') return null;
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : null;
  };
  const initialResolvedJobsheetId = resolveJobsheetId(targetJobsheetId !== undefined ? targetJobsheetId : initialJobsheetId);
  const numericBusinessId = Number(businessId) || 0;
  const [business, setBusiness] = useState(businessName ? { id: numericBusinessId, business_name: businessName } : null);
  const [formState, setFormState] = useState(DEFAULT_JOBSHEET(numericBusinessId));
  const [jobsheetId, setJobsheetId] = useState(initialResolvedJobsheetId);
  const [venues, setVenues] = useState([]);
  const [fieldGroups, setFieldGroups] = useState(FALLBACK_JOBSHEET_GROUPS);
  const [pricingConfig, setPricingConfig] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [venueSaving, setVenueSaving] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const [documentGeneratingKey, setDocumentGeneratingKey] = useState(null);
  const [lastOutputPath, setLastOutputPath] = useState('');
  const [jobsheetDocuments, setJobsheetDocuments] = useState([]);
  const [jobsheetDocumentsFolder, setJobsheetDocumentsFolder] = useState('');
  const [jobsheetDocumentsLoading, setJobsheetDocumentsLoading] = useState(false);
  const [jobsheetDocumentsError, setJobsheetDocumentsError] = useState('');
  const [documentDefinitions, setDocumentDefinitions] = useState([]);
  const [documentDefinitionsLoading, setDocumentDefinitionsLoading] = useState(false);
  const [documentDefinitionsError, setDocumentDefinitionsError] = useState('');
  const [selectedDefinitionKey, setSelectedDefinitionKey] = useState(null);
  const [definitionModalOpen, setDefinitionModalOpen] = useState(false);
  const [definitionModalMode, setDefinitionModalMode] = useState('create');
  const [definitionDraft, setDefinitionDraft] = useState(createDefinitionDraft());
  const [definitionModalError, setDefinitionModalError] = useState('');
  const [definitionKeyEdited, setDefinitionKeyEdited] = useState(false);
  const [definitionSaving, setDefinitionSaving] = useState(false);
  const formStateRef = useRef(DEFAULT_JOBSHEET(numericBusinessId));
  const [activeEditorSection, setActiveEditorSection] = useState('client');
  const sectionRestoredRef = useRef(false);

  const autoSaveTimer = useRef(null);
  const initialLoadRef = useRef(true);
  const creatingRef = useRef(false);
  const createDraftTimerRef = useRef(null);
  const previousJobsheetIdRef = useRef(initialResolvedJobsheetId);
  const nameDateRef = useRef({ name: '', date: '' });

  const storagePrefix = useMemo(() => (
    `jobsheetEditor:${isInline ? 'inline' : 'window'}:${numericBusinessId}:`
  ), [isInline, numericBusinessId]);

  const getStoredSection = useCallback((jobsheetValue) => {
    const id = jobsheetValue != null ? Number(jobsheetValue) : null;
    if (!id || Number.isNaN(id)) return null;
    try {
      const sessionVal = window.sessionStorage.getItem(`${storagePrefix}${id}`) || null;
      const persistFlag = window.localStorage.getItem('app:persistUiState') === 'true';
      if (sessionVal) return sessionVal;
      if (persistFlag) {
        return window.localStorage.getItem(`${storagePrefix}${id}`) || null;
      }
      return null;
    } catch (_err) {
      return null;
    }
  }, [storagePrefix]);

  const storeSection = useCallback((jobsheetValue, section) => {
    const id = jobsheetValue != null ? Number(jobsheetValue) : null;
    if (!id || Number.isNaN(id)) return;
    if (!section) return;
    try {
      window.sessionStorage.setItem(`${storagePrefix}${id}`, section);
      if (window.localStorage.getItem('app:persistUiState') === 'true') {
        window.localStorage.setItem(`${storagePrefix}${id}`, section);
      }
    } catch (_err) {
      // ignore storage errors
    }
  }, [storagePrefix]);

  useEffect(() => {
    previousJobsheetIdRef.current = jobsheetId != null ? Number(jobsheetId) : null;
  }, [jobsheetId]);

  useEffect(() => {
    const id = jobsheetId != null ? Number(jobsheetId) : null;
    if (!id || !activeEditorSection) return;
    // Avoid writing a default tab before we restore the saved one
    if (!sectionRestoredRef.current) return;
    storeSection(id, activeEditorSection);
  }, [jobsheetId, activeEditorSection, storeSection]);

  // Track client name/date for folder and filenames; initialize on load
  useEffect(() => {
    const nm = String(formState?.client_name || '');
    const dt = String(formState?.event_date || '');
    nameDateRef.current = { name: nm, date: dt };
  }, [jobsheetId]);

  // Note: restoration of the active editor section is handled on mount
  // via getStoredSection when a jobsheet is loaded; no dependency on
  // workspaceSection here to avoid undefined references in this window.

  const loadDocumentDefinitions = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDocumentDefinitions([]);
      setDocumentDefinitionsLoading(false);
      setDocumentDefinitionsError('');
      return;
    }
    setDocumentDefinitionsLoading(true);
    setDocumentDefinitionsError('');
    try {
      const api = window.api;
      if (!api || typeof api.getDocumentDefinitions !== 'function') {
        throw new Error('Unable to load document definitions: API unavailable');
      }
      const data = await api.getDocumentDefinitions(numericBusinessId, { includeInactive: true });
      const list = Array.isArray(data) ? data.map(def => ({ ...def })) : [];
      setDocumentDefinitions(list);

      if (!list.length) {
        setSelectedDefinitionKey(null);
        return;
      }

      const hasSelection = list.some(def => def.key === selectedDefinitionKey);
      if (!hasSelection) {
        const fallback = list[0];
        setSelectedDefinitionKey(fallback ? fallback.key : null);
      }
    } catch (err) {
      console.error('Failed to load document definitions', err);
      setDocumentDefinitions([]);
      setDocumentDefinitionsError(err?.message || 'Unable to load document definitions');
    } finally {
      setDocumentDefinitionsLoading(false);
    }
  }, [numericBusinessId, selectedDefinitionKey]);

  const selectDefinitionKey = useCallback((key) => {
    setDocumentDefinitionsError('');
    setSelectedDefinitionKey(key);
  }, []);

  useEffect(() => {
    loadDocumentDefinitions();
  }, [loadDocumentDefinitions]);

  const refreshJobsheetDocuments = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocuments([]);
      setJobsheetDocumentsFolder('');
      setJobsheetDocumentsLoading(false);
      setJobsheetDocumentsError('');
      return;
    }
    if (!jobsheetId) {
      setJobsheetDocuments([]);
      setJobsheetDocumentsFolder('');
      setJobsheetDocumentsLoading(false);
      setJobsheetDocumentsError('');
      return;
    }
    setJobsheetDocumentsLoading(true);
    setJobsheetDocumentsError('');
    try {
      const api = window.api;
      if (!api || typeof api.listJobsheetDocuments !== 'function') {
        throw new Error('Unable to load documents: API unavailable');
      }

      const normalizedJobsheetId = Number(jobsheetId);
      const currentState = formStateRef.current || DEFAULT_JOBSHEET(numericBusinessId);

      const fetchDocuments = async () => {
        const response = await api.listJobsheetDocuments({
          businessId: numericBusinessId,
          jobsheetId: normalizedJobsheetId,
          jobsheetSnapshot: currentState
        });
        setJobsheetDocumentsFolder(response?.jobsheet_folder || '');
        return Array.isArray(response?.documents) ? response.documents : [];
      };

      const filterForJobsheet = (docs) => {
        return docs.filter(doc => {
          const docJobsheetId = doc?.jobsheet_id != null ? Number(doc.jobsheet_id) : null;
          if (docJobsheetId != null && docJobsheetId === normalizedJobsheetId) {
            return true;
          }
          if (docJobsheetId != null && docJobsheetId !== normalizedJobsheetId) {
            return false;
          }
          return matchesDocumentToJobsheet(doc, currentState);
        });
      };

      let documentsList = await fetchDocuments();
      let filtered = filterForJobsheet(documentsList);

      if (typeof api.syncJobsheetOutputs === 'function') {
        try {
          const syncResult = await api.syncJobsheetOutputs({
            businessId: numericBusinessId,
            jobsheetId: normalizedJobsheetId,
            jobsheetSnapshot: currentState,
            hintPaths: filtered.map(doc => doc?.file_path).filter(Boolean)
          });

          if (syncResult?.added > 0) {
            documentsList = await fetchDocuments();
            filtered = filterForJobsheet(documentsList);

            const newIds = Array.isArray(syncResult.records)
              ? syncResult.records.map(item => item?.document_id).filter(id => id != null)
              : [];

            if (newIds.length) {
              window.api?.notifyJobsheetChange?.({
                type: 'documents-updated',
                businessId: numericBusinessId,
                jobsheetId: normalizedJobsheetId,
                documentIds: newIds
              });
            }
          }
        } catch (syncErr) {
          console.error('Failed to sync exported documents', syncErr);
        }
      }

      setJobsheetDocuments(filtered);
    } catch (err) {
      console.error('Failed to load jobsheet documents', err);
      setJobsheetDocumentsError(err?.message || 'Unable to load documents');
      setJobsheetDocuments([]);
      setJobsheetDocumentsFolder('');
    } finally {
      setJobsheetDocumentsLoading(false);
    }
  }, [jobsheetId, numericBusinessId]);

  useEffect(() => {
    if (!jobsheetId) {
      setJobsheetDocuments([]);
      setJobsheetDocumentsError('');
      setJobsheetDocumentsLoading(false);
      setJobsheetDocumentsFolder('');
      return;
    }
    refreshJobsheetDocuments();
  }, [jobsheetId, refreshJobsheetDocuments]);

  useEffect(() => {
    if (!DOCUMENT_FEATURES_ENABLED && !DOCUMENT_GENERATION_ENABLED) return () => {};
    if (!window.api) return () => {};
    window.api.watchDocuments?.({ businessId: numericBusinessId }).catch(err => {
      console.warn('Unable to start jobsheet documents watcher', err);
    });
    const unsubscribe = window.api.onDocumentsChange?.((payload) => {
      if (!payload || payload.businessId !== numericBusinessId) return;
      refreshJobsheetDocuments();
    });
    return () => {
      unsubscribe?.();
    };
  }, [numericBusinessId, refreshJobsheetDocuments]);

  const findDefinitionByKey = useCallback((key) => {
    if (!key) return null;
    return documentDefinitions.find(definition => definition.key === key) || null;
  }, [documentDefinitions]);

  const activeDocumentDefinition = useMemo(() => findDefinitionByKey(selectedDefinitionKey), [selectedDefinitionKey, findDefinitionByKey]);


  useEffect(() => {
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload || payload.businessId !== numericBusinessId) return;
      if (payload.type !== 'documents-updated') return;
      if (!jobsheetId) return;
      const payloadJobsheetId = payload.jobsheetId != null
        ? Number(payload.jobsheetId)
        : payload.document?.jobsheet_id != null
          ? Number(payload.document.jobsheet_id)
          : null;
      if (payloadJobsheetId == null || payloadJobsheetId === Number(jobsheetId)) {
        refreshJobsheetDocuments();
      }
    });
    return () => unsubscribe();
  }, [jobsheetId, numericBusinessId, refreshJobsheetDocuments]);

  const handleOpenDocumentFile = useCallback(async (filePath) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document access is currently disabled.');
      return;
    }
    if (!filePath) {
      setJobsheetDocumentsError('Document file not available');
      return;
    }
    try {
      setJobsheetDocumentsError('');
      const response = await window.api?.openPath?.(filePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open document');
      }
    } catch (err) {
      console.error('Failed to open document', err);
      setJobsheetDocumentsError(err?.message || 'Unable to open document');
    }
  }, []);

  const handleRevealDocument = useCallback(async (filePath) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document access is currently disabled.');
      return;
    }
    if (!filePath) {
      setJobsheetDocumentsError('Document file not available');
      return;
    }
    try {
      setJobsheetDocumentsError('');
      const result = await window.api?.showItemInFolder?.(filePath);
      if (result && result.ok === false) {
        throw new Error(result.message || 'Unable to reveal document');
      }
    } catch (err) {
      console.error('Failed to reveal document', err);
      setJobsheetDocumentsError(err?.message || 'Unable to locate document on disk');
    }
  }, []);

  const handleExportWorkbookPdf = useCallback(async (doc, options = {}) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document export is currently disabled.');
      return;
    }
    if (!doc || !doc.file_path) {
      setJobsheetDocumentsError('Workbook file not available for export');
      return;
    }
    try {
      setJobsheetDocumentsError('');
      const payload = { businessId: numericBusinessId, filePath: doc.file_path };
      if (options && Number.isInteger(options.requestedNumber) && options.requestedNumber > 0) {
        payload.requestedNumber = Number(options.requestedNumber);
      }
      const result = await window.api?.exportWorkbookPdfs?.(payload);
      if (result && result.ok === false) {
        throw new Error(result.message || 'Unable to export workbook to PDF');
      }

      if (Array.isArray(result?.outputs)) {
        const successes = result.outputs.filter(item => item && item.success && item.file_path);
        if (successes.length) {
          const firstPath = successes[0].file_path;
          if (firstPath) {
            setLastOutputPath(firstPath);
          }
          const labels = successes.map(item => item.label || item.sheet || 'PDF').join(', ');
          setMessage(`Exported ${labels}`);
          setTimeout(() => setMessage(''), 2500);
        }
      }

      await refreshJobsheetDocuments();

      window.api?.notifyJobsheetChange?.({
        type: 'documents-updated',
        businessId: numericBusinessId,
        jobsheetId: jobsheetId != null ? Number(jobsheetId) : null
      });
      return result;
    } catch (err) {
      console.error('Failed to export workbook PDFs', err);
      setJobsheetDocumentsError(err?.message || 'Unable to export workbook to PDF');
      return { ok: false, message: err?.message || 'Unable to export workbook to PDF' };
    }
  }, [jobsheetId, numericBusinessId, refreshJobsheetDocuments, setMessage]);

  

  const handleDeleteJobsheetDocument = useCallback(async (doc) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document access is currently disabled.');
      return;
    }
    if (!doc || doc.document_id == null) return;
    // Delete silently and also remove the generated file from disk
    const removeFile = true;

    try {
      setJobsheetDocumentsError('');
      if (doc.is_locked) {
        const proceed = window.confirm('This document is locked. Unlock and delete it (including the file on disk)?');
        if (!proceed) return;
        try { await window.api?.setDocumentLock?.(doc.document_id, false); } catch (_) {}
      }
      await window.api?.deleteDocument?.(doc.document_id, { removeFile });
      setMessage('Document deleted');
      await refreshJobsheetDocuments();
      window.api?.notifyJobsheetChange?.({
        type: 'documents-updated',
        businessId: numericBusinessId,
        jobsheetId,
        documentId: doc.document_id
      });
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to delete document', err);
      setJobsheetDocumentsError(err?.message || 'Unable to delete document');
    }
  }, [jobsheetId, numericBusinessId, refreshJobsheetDocuments, setMessage]);

  const openNewDefinitionModal = useCallback(() => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) return;
    setDefinitionModalMode('create');
    setDefinitionDraft(createDefinitionDraft());
    setDefinitionModalError('');
    setDefinitionKeyEdited(false);
    setDocumentDefinitionsError('');
    setDefinitionModalOpen(true);
  }, []);

  const openEditDefinitionModal = useCallback((definition) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) return;
    if (!definition) return;
    setDefinitionModalMode('edit');
    setDefinitionDraft(createDefinitionDraft({
      ...definition,
      template_path: definition.template_path || '',
      is_primary: definition.is_primary ? 1 : 0,
      is_active: definition.is_active === 0 ? 0 : 1,
      is_locked: definition.is_locked ? 1 : 0,
      sort_order: definition.sort_order != null ? definition.sort_order : null
    }));
    setDefinitionModalError('');
    setDefinitionKeyEdited(true);
    setDocumentDefinitionsError('');
    setDefinitionModalOpen(true);
  }, []);

  const handleCloseDefinitionModal = useCallback(() => {
    setDefinitionModalOpen(false);
    setDefinitionSaving(false);
    setDefinitionModalError('');
    setDefinitionDraft(createDefinitionDraft());
    setDefinitionKeyEdited(false);
    setDefinitionModalMode('create');
  }, []);

  const handleDefinitionDraftChange = useCallback((field, rawValue) => {
    setDefinitionDraft(prev => {
      const next = { ...prev };
      const value = rawValue;

      switch (field) {
        case 'label':
          next.label = value;
          if (!definitionKeyEdited) {
            next.key = slugifyDefinitionKey(value);
          }
          break;
        case 'key':
          next.key = slugifyDefinitionKey(value);
          break;
        case 'doc_type':
          next.doc_type = value;
          if (value !== 'invoice') {
            next.invoice_variant = '';
          }
          break;
        case 'invoice_variant':
          next.invoice_variant = value;
          break;
        case 'description':
          next.description = value;
          break;
        case 'template_path':
          next.template_path = value || '';
          break;
        case 'is_primary':
          next.is_primary = value ? 1 : 0;
          break;
        case 'is_active':
          next.is_active = value ? 1 : 0;
          break;
        default:
          next[field] = value;
          break;
      }

      return next;
    });

    if (field === 'key') {
      setDefinitionKeyEdited(true);
    }
    setDefinitionModalError('');
  }, [definitionKeyEdited]);

  const handlePickDefinitionDraftTemplate = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDefinitionModalError('Document generation features are disabled.');
      return;
    }
    const api = window.api;
    if (!api || typeof api.chooseFile !== 'function') {
      setDefinitionModalError('Unable to select template: API unavailable');
      return;
    }

    const meta = DOC_TYPE_META[definitionDraft.doc_type] || null;
    try {
      const selectedPath = await api.chooseFile({
        title: `Choose template for ${definitionDraft.label || meta?.label || 'document'}`,
        defaultPath: definitionDraft.template_path || undefined,
        filters: meta?.filters
      });
      if (!selectedPath) return;
      handleDefinitionDraftChange('template_path', selectedPath);
      setDefinitionModalError('');
    } catch (err) {
      console.error('Failed to choose template file', err);
      setDefinitionModalError(err?.message || 'Unable to choose template file');
    }
  }, [definitionDraft, handleDefinitionDraftChange]);

  const handleClearDefinitionDraftTemplate = useCallback(() => {
    handleDefinitionDraftChange('template_path', '');
  }, [handleDefinitionDraftChange]);

  const handleOpenDraftTemplate = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDefinitionModalError('Document generation features are disabled.');
      return;
    }
    const templatePath = definitionDraft.template_path;
    if (!templatePath) {
      setDefinitionModalError('Select a template before opening it.');
      return;
    }
    try {
      const response = await window.api?.openPath?.(templatePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open template');
      }
    } catch (err) {
      console.error('Failed to open template', err);
      setDefinitionModalError(err?.message || 'Unable to open template');
    }
  }, [definitionDraft.template_path]);

  const handleNormalizeDraftTemplate = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDefinitionModalError('Document generation features are disabled.');
      return;
    }
    const templatePath = definitionDraft.template_path;
    if (!templatePath) {
      setDefinitionModalError('Select a template before normalizing it.');
      return;
    }
    try {
      const response = await window.api?.normalizeTemplate?.({ templatePath });
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to normalize template');
      }
      setDefinitionModalError('');
      setMessage('Template normalized');
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to normalize template', err);
      setDefinitionModalError(err?.message || 'Unable to normalize template');
    }
  }, [definitionDraft.template_path, setMessage]);

  const handleSaveDefinition = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDefinitionModalError('Document generation features are disabled.');
      return;
    }
    const api = window.api;
    if (!api || typeof api.saveDocumentDefinition !== 'function') {
      setDefinitionModalError('Unable to save definition: API unavailable');
      return;
    }

    const trimmedLabel = (definitionDraft.label || '').trim();
    const trimmedKey = slugifyDefinitionKey(definitionDraft.key);
    const docType = (definitionDraft.doc_type || '').trim();

    if (!trimmedLabel) {
      setDefinitionModalError('Label is required');
      return;
    }
    if (!trimmedKey) {
      setDefinitionModalError('Key is required');
      return;
    }
    if (!DOC_TYPE_META[docType]) {
      setDefinitionModalError('Choose a valid document type');
      return;
    }

    const payload = {
      key: trimmedKey,
      label: trimmedLabel,
      doc_type: docType,
      description: definitionDraft.description ? String(definitionDraft.description) : null,
      invoice_variant: docType === 'invoice' && definitionDraft.invoice_variant ? String(definitionDraft.invoice_variant) : null,
      template_path: definitionDraft.template_path ? String(definitionDraft.template_path) : null,
      // requires_total removed
      is_primary: 0,
      is_active: definitionDraft.is_active === 0 ? 0 : 1,
      is_locked: definitionDraft.is_locked ? 1 : 0,
      sort_order: definitionDraft.sort_order != null ? Number(definitionDraft.sort_order) : null
    };

    setDefinitionModalError('');
    setDefinitionSaving(true);
    try {
      await api.saveDocumentDefinition(numericBusinessId, payload);

      if (payload.template_path && DOC_TYPE_META[payload.doc_type]?.supportsNormalize) {
        try {
          await window.api?.normalizeTemplate?.({ templatePath: payload.template_path });
        } catch (normalizeErr) {
          console.warn('Failed to normalize template', normalizeErr);
        }
      }

      await loadDocumentDefinitions();
      selectDefinitionKey(trimmedKey);
      handleCloseDefinitionModal();
    } catch (err) {
      console.error('Failed to save document definition', err);
      setDefinitionModalError(err?.message || 'Unable to save document definition');
    } finally {
      setDefinitionSaving(false);
    }
  }, [definitionDraft, numericBusinessId, loadDocumentDefinitions, handleCloseDefinitionModal, selectDefinitionKey]);

  const handleDeleteDefinition = useCallback(async (definition) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setDocumentDefinitionsError('Document generation features are disabled.');
      return;
    }
    if (!definition) return;
    if (definition.is_locked) {
      setDefinitionModalError('This definition is locked and cannot be deleted.');
      return;
    }
    const confirmed = window.confirm(`Delete ${definition.label || definition.key}? This cannot be undone.`);
    if (!confirmed) return;
    const api = window.api;
    if (!api || typeof api.deleteDocumentDefinition !== 'function') {
      setDefinitionModalError('Unable to delete definition: API unavailable');
      return;
    }
    try {
      await api.deleteDocumentDefinition(numericBusinessId, definition.key);
      setMessage('Document definition deleted');
      setTimeout(() => setMessage(''), 1500);
      await loadDocumentDefinitions();
      if (selectedDefinitionKey === definition.key) {
        selectDefinitionKey(null);
      }
      handleCloseDefinitionModal();
    } catch (err) {
      console.error('Failed to delete document definition', err);
      setDefinitionModalError(err?.message || 'Unable to delete document definition');
    }
  }, [numericBusinessId, loadDocumentDefinitions, selectedDefinitionKey, selectDefinitionKey, handleCloseDefinitionModal, setMessage]);

  const handleOpenDefinitionTemplate = useCallback(async (definition) => {
    if (!definition) return;
    const templatePath = definition.template_path;
    if (!templatePath) {
      setDocumentDefinitionsError('No template configured for this document type. Set one before opening.');
      return;
    }
    try {
      const response = await window.api?.openPath?.(templatePath);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open template');
      }
    } catch (err) {
      console.error('Failed to open definition template', err);
      setDocumentDefinitionsError(err?.message || 'Unable to open definition template');
    }
  }, []);

  const definitionModalTitle = definitionModalMode === 'edit'
    ? 'Edit document type'
    : 'New document type';
  const modalDocMeta = DOC_TYPE_META[definitionDraft.doc_type] || {};
  const modalTemplatePath = definitionDraft.template_path || '';
  const modalHasTemplate = modalTemplatePath !== '';
  const modalSupportsNormalize = modalDocMeta.supportsNormalize && modalHasTemplate && modalTemplatePath.toLowerCase().endsWith('.xlsx');
  const modalIsLocked = definitionModalMode === 'edit' && Boolean(definitionDraft.is_locked);

  useEffect(() => {
    if (!jobsheetId) return;
    window.api?.notifyJobsheetChange?.({
      type: 'jobsheet-editor-focus',
      businessId: numericBusinessId,
      jobsheetId,
      active: true
    });
    return () => {
      window.api?.notifyJobsheetChange?.({
        type: 'jobsheet-editor-focus',
        businessId: numericBusinessId,
        jobsheetId,
        active: false
      });
    };
  }, [numericBusinessId, jobsheetId, setBusiness, setFormState, setLoading, setActiveEditorSection, setError, setMessage]);

  useEffect(() => {
    if (!isInline) return;
    if (targetJobsheetId === undefined) return;
    const normalize = (value) => {
      if (value === undefined || value === null) return null;
      if (value === '' || value === 'new') return null;
      const numericValue = Number(value);
      return Number.isFinite(numericValue) ? numericValue : null;
    };
    const nextTarget = normalize(targetJobsheetId);
    const current = normalize(jobsheetId);
    if (nextTarget === current) return;

    initialLoadRef.current = true;
    creatingRef.current = false;
    setError('');
    setMessage('');
    if (nextTarget != null) {
      const storedSection = getStoredSection(nextTarget);
      const fallbackSection = storedSection || 'client';
      setActiveEditorSection(fallbackSection);
      sectionRestoredRef.current = true;
    } else {
      setActiveEditorSection('client');
      sectionRestoredRef.current = true;
    }
    setLastOutputPath('');

    if (nextTarget != null) {
      setLoading(true);
      setJobsheetId(nextTarget);
    } else {
      const resetState = DEFAULT_JOBSHEET(numericBusinessId);
      setJobsheetId(null);
      setFormState(resetState);
      formStateRef.current = resetState;
      setLoading(false);
    }
  }, [isInline, targetJobsheetId, jobsheetId, numericBusinessId, getStoredSection]);

  useEffect(() => {
    if (isInline) return () => {};
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload || payload.businessId !== numericBusinessId) return;
      if (payload.type !== 'jobsheet-load-request') return;
      const requestedId = payload.jobsheetId != null ? Number(payload.jobsheetId) : null;
      if (requestedId != null && requestedId === jobsheetId) {
        window.focus();
        return;
      }
      initialLoadRef.current = true;
      creatingRef.current = false;
      setError('');
      setMessage('');
      if (requestedId != null) {
        const storedSection = getStoredSection(requestedId);
        const fallbackSection = storedSection || 'client';
        setActiveEditorSection(fallbackSection);
        sectionRestoredRef.current = true;
      } else {
        setActiveEditorSection('client');
        sectionRestoredRef.current = true;
      }
      if (payload.businessName) {
        setBusiness(prev => prev || { id: numericBusinessId, business_name: payload.businessName });
      }
      if (requestedId != null) {
        setLoading(true);
        setJobsheetId(requestedId);
        const url = new URL(window.location.href);
        url.searchParams.set('jobsheetId', requestedId);
        window.history.replaceState({}, '', url.toString());
      } else {
        const resetState = DEFAULT_JOBSHEET(numericBusinessId);
        setJobsheetId(null);
        setFormState(resetState);
        formStateRef.current = resetState;
        setLoading(false);
        const url = new URL(window.location.href);
        url.searchParams.set('jobsheetId', 'new');
        window.history.replaceState({}, '', url.toString());
      }
    });
    return () => unsubscribe();
  }, [isInline, numericBusinessId, jobsheetId, getStoredSection]);

  const buildSnapshot = useCallback((state, id) => ({
    jobsheet_id: id ?? state.jobsheet_id ?? null,
    client_name: state.client_name || '',
    event_type: state.event_type || '',
    event_date: state.event_date || '',
    venue_name: state.venue_name || '',
    venue_town: state.venue_town || '',
    status: state.status || 'enquiry',
    ahmen_fee: state.ahmen_fee !== undefined && state.ahmen_fee !== null && state.ahmen_fee !== ''
      ? Number(state.ahmen_fee)
      : null,
    pricing_total: state.pricing_total !== undefined && state.pricing_total !== null && state.pricing_total !== ''
      ? Number(state.pricing_total)
      : null,
    updated_at: new Date().toISOString()
  }), []);

  useEffect(() => {
    let mounted = true;
    const load = async () => {
      const api = window.api;
      if (!api || !api.businessSettings) {
        setError('Application API unavailable in editor window');
        setLoading(false);
        return;
      }
      if (!numericBusinessId) {
        setError('Missing business reference');
        setLoading(false);
        return;
      }
      setLoading(true);
      try {
        const mergeFieldPromise = typeof api.getMergeFields === 'function'
          ? api.getMergeFields()
          : Promise.resolve([]);
        const [businessList, venueData, pricingData, mergeFieldData] = await Promise.all([
          api.businessSettings(),
          api.getAhmenVenues({ businessId: numericBusinessId }),
          api.getAhmenPricing(),
          mergeFieldPromise
        ]);
        if (!mounted) return;
        const foundBusiness = (businessList || []).find(item => item.id === numericBusinessId) || null;
        if (foundBusiness) {
          setBusiness(prev => {
            if (prev && prev.id === foundBusiness.id && prev.save_path === foundBusiness.save_path && prev.business_name === foundBusiness.business_name) {
              return prev;
            }
            return { ...foundBusiness };
          });
        }
        setVenues(normalizeVenues(venueData));
        setPricingConfig(pricingData || null);
        setFieldGroups(buildJobsheetGroups(mergeFieldData || []));

        let effectiveJobsheetId = jobsheetId;
        if (!effectiveJobsheetId && !isInline && initialJobsheetId && initialJobsheetId !== 'new') {
          effectiveJobsheetId = Number(initialJobsheetId);
          setJobsheetId(effectiveJobsheetId);
        }

        if (effectiveJobsheetId) {
          const sheet = await api.getAhmenJobsheet(effectiveJobsheetId);
          if (!mounted) return;
          if (sheet) {
            setFormState(mapApiToForm(sheet, numericBusinessId));
            const storedSection = getStoredSection(effectiveJobsheetId);
            if (storedSection) {
              setActiveEditorSection(storedSection);
            } else if (initialLoadRef.current) {
              setActiveEditorSection(prev => prev || 'client');
            }
            sectionRestoredRef.current = true;
          }
        } else {
          setFormState(DEFAULT_JOBSHEET(numericBusinessId));
          setActiveEditorSection('client');
          sectionRestoredRef.current = true;
        }
        initialLoadRef.current = false;
      } catch (err) {
        if (!mounted) return;
        console.error('Failed to load jobsheet editor', err);
        setError(err?.message || 'Unable to load jobsheet');
      } finally {
        if (mounted) setLoading(false);
      }
    };
    load();
    return () => {
      mounted = false;
    };
  }, [numericBusinessId, initialJobsheetId, jobsheetId, isInline, getStoredSection]);

  useEffect(() => {
    formStateRef.current = formState;
  }, [formState]);

  const handleUpdateSingerPool = useCallback(async (singers) => {
    const api = window.api;
    if (!api || !api.updateAhmenSingerPool) {
      throw new Error('Unable to update singer pool: API unavailable');
    }
    const nextConfig = await api.updateAhmenSingerPool(singers);
    setPricingConfig(nextConfig || null);
    setActiveEditorSection('pricing');
    return nextConfig;
  }, []);

  const pricingDerived = useMemo(() => {
    if (!pricingConfig) return null;
    const pool = Array.isArray(pricingConfig.singerPool) ? pricingConfig.singerPool : [];
    const poolMap = new Map(pool.map(singer => [String(singer.id), singer]));
    const selectedEntries = normalizeSingerEntries(formState.pricing_selected_singers);
    let base = 0;
    let singerCount = 0;

    selectedEntries.forEach(entry => {
      const singer = poolMap.get(entry.id);
      let feeValue = 0;
      if (singer) {
        feeValue = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
          ? Number(entry.fee)
          : Number(singer.fee);
        singerCount += 1;
      } else if (entry.custom) {
        feeValue = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
          ? Number(entry.fee)
          : 0;
        singerCount += 1;
      } else {
        return;
      }
      base += Number.isFinite(feeValue) ? feeValue : 0;
    });

    const custom = Number(formState.pricing_custom_fees) || 0;
    const singerSubtotal = base + custom;
    const singerDiscountValue = calculateDiscountValue({
      type: formState.pricing_discount_type || 'amount',
      value: formState.pricing_discount,
      subtotal: singerSubtotal
    });
    const singerNet = Math.max(singerSubtotal - singerDiscountValue, 0);

    const productionSubtotal = Number(formState.pricing_production_subtotal) || 0;
    const productionDiscountValue = calculateDiscountValue({
      type: formState.pricing_production_discount_type || 'amount',
      value: formState.pricing_production_discount,
      subtotal: productionSubtotal
    });
    const productionNet = Math.max(productionSubtotal - productionDiscountValue, 0);

    const total = Math.max(singerNet + productionNet, 0);
    const hasSelection = singerCount > 0 || custom !== 0 || productionSubtotal !== 0 || singerDiscountValue > 0 || productionDiscountValue > 0;
    const totalString = hasSelection ? total.toFixed(2) : '';
    return {
      base,
      singerCount,
      custom,
      singerSubtotal,
      singerNet,
      singerDiscountValue,
      productionSubtotal,
      productionNet,
      productionDiscountValue,
      hasSelection,
      total,
      totalString
    };
  }, [pricingConfig, formState]);

  useEffect(() => {
    if (!pricingDerived) return;
    setFormState(prev => {
      const nextTotal = pricingDerived.totalString || '';
      const currentTotal = prev.pricing_total ?? '';
      const shouldUpdateTotal = nextTotal !== currentTotal;

      let shouldUpdateFee = false;
      let nextFeeValue = prev.ahmen_fee ?? '';
      const derivedAhmenFee = pricingDerived.hasSelection ? pricingDerived.singerNet.toFixed(2) : '';
      if (pricingDerived.hasSelection) {
        if (derivedAhmenFee && derivedAhmenFee !== (prev.ahmen_fee ?? '')) {
          shouldUpdateFee = true;
          nextFeeValue = derivedAhmenFee;
        }
      } else if (!pricingDerived.hasSelection && !pricingDerived.totalString && prev.ahmen_fee) {
        shouldUpdateFee = true;
        nextFeeValue = '';
      }

      if (!shouldUpdateTotal && !shouldUpdateFee) return prev;

      const next = { ...prev };
      if (shouldUpdateTotal) next.pricing_total = nextTotal;
      if (shouldUpdateFee) next.ahmen_fee = nextFeeValue;
      return applyDerivedFields(next);
    });
  }, [pricingDerived]);

  useEffect(() => {
    setFormState(prev => {
      const nextProduction = formState.pricing_production_total ?? '';
      const currentFees = prev.production_fees ?? '';
      const previousAuto = prev.pricing_production_total ?? '';
      const shouldUpdate = currentFees === previousAuto || currentFees === '';
      if (!shouldUpdate) return prev;
      if (currentFees === nextProduction) return prev;
      return applyDerivedFields({ ...prev, production_fees: nextProduction });
    });
  }, [formState.pricing_production_total]);

  useEffect(() => {
    if (loading || !jobsheetId) return;
    if (initialLoadRef.current) {
      initialLoadRef.current = false;
      return;
    }
    if (autoSaveTimer.current) clearTimeout(autoSaveTimer.current);
    const api = window.api;
    if (!api || !api.updateAhmenJobsheet) return;
    // Skip autosave until a client name is entered to avoid backend validation errors
    const hasClientName = !!String(formState?.client_name || '').trim();
    if (!hasClientName) {
      // Clear any previous autosave error to keep UI calm while drafting
      // setError(''); // optional: keep silent
      return;
    }
    autoSaveTimer.current = setTimeout(async () => {
      setSaving(true);
      try {
        const payload = preparePayload(formState, numericBusinessId);
        await api.updateAhmenJobsheet(jobsheetId, payload);
        // If name or date changed, rename folder and filenames
        try {
          const prev = nameDateRef.current || { name: '', date: '' };
          const currentName = String(formState?.client_name || '');
          const currentDate = String(formState?.event_date || '');
          if (prev.name !== currentName || prev.date !== currentDate) {
            await window.api?.renameJobsheetArtifacts?.({ businessId: numericBusinessId, jobsheetId });
            nameDateRef.current = { name: currentName, date: currentDate };
          }
        } catch (_) {}
        setMessage('Saved');
        window.api?.notifyJobsheetChange?.({
          type: 'jobsheet-updated',
          businessId: numericBusinessId,
          jobsheetId,
          snapshot: buildSnapshot(formState, jobsheetId)
        });
        setTimeout(() => setMessage(''), 1500);
      } catch (err) {
        console.error('Failed to auto-save jobsheet', err);
        setError(err?.message || 'Unable to save jobsheet');
      } finally {
        setSaving(false);
      }
    }, 600);
    return () => clearTimeout(autoSaveTimer.current);
  }, [formState, jobsheetId, numericBusinessId, loading]);

  const saveJobsheet = useCallback(async () => {
    if (loading || !jobsheetId) return;
    const api = window.api;
    if (!api || !api.updateAhmenJobsheet) return;
    const currentState = formStateRef.current;
    const hasClientName = !!String(currentState?.client_name || '').trim();
    if (!hasClientName) {
      setError('Enter a client name before saving');
      return;
    }
    setSaving(true);
    try {
      const payload = preparePayload(currentState, numericBusinessId);
      await api.updateAhmenJobsheet(jobsheetId, payload);
      try {
        const prev = nameDateRef.current || { name: '', date: '' };
        const currentName = String(currentState?.client_name || '');
        const currentDate = String(currentState?.event_date || '');
        if (prev.name !== currentName || prev.date !== currentDate) {
          await window.api?.renameJobsheetArtifacts?.({ businessId: numericBusinessId, jobsheetId });
          nameDateRef.current = { name: currentName, date: currentDate };
        }
      } catch (_) {}
      window.api?.notifyJobsheetChange?.({
        type: 'jobsheet-updated',
        businessId: numericBusinessId,
        jobsheetId,
        snapshot: buildSnapshot(currentState, jobsheetId)
      });
      setMessage('Saved');
      setTimeout(() => setMessage(''), 1200);
    } catch (err) {
      console.error('Failed to auto-save jobsheet', err);
      setError(err?.message || 'Unable to save jobsheet');
    } finally {
      setSaving(false);
    }
  }, [buildSnapshot, jobsheetId, numericBusinessId, loading]);

  // Debounced autosave handles persistence; avoid immediate save loop here

  const resolvedBusiness = useMemo(() => (
    business ? { ...business } : {
      id: numericBusinessId,
      business_name: businessName || 'Jobsheet',
      save_path: business?.save_path || ''
    }
  ), [business, numericBusinessId, businessName]);

  const parseAmount = useCallback((value) => {
    if (value === null || value === undefined || value === '') return null;
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return null;
    return Math.round(numeric * 100) / 100;
  }, []);

  const buildDocumentPayload = useCallback((definition) => {
    if (!definition) return null;
    const docType = (definition.doc_type || '').toLowerCase();
    if (!docType) return null;

    const current = formStateRef.current || DEFAULT_JOBSHEET(numericBusinessId);

    const productionItems = normalizeProductionItems(current.pricing_production_items);
    const productionSubtotal = parseAmount(current.pricing_production_subtotal) ?? productionItems.reduce((sum, item) => sum + (parseAmount(item.cost) || 0), 0);

    const totalAmount = parseAmount(current.pricing_total)
      ?? (pricingDerived ? parseAmount(pricingDerived.total) : null)
      ?? parseAmount(current.ahmen_fee);

    const depositAmount = parseAmount(current.deposit_amount);
    const balanceAmount = parseAmount(current.balance_amount);
    const extraFees = parseAmount(current.extra_fees ?? current.pricing_custom_fees);
    const productionFees = parseAmount(current.production_fees) ?? productionSubtotal;

    const discountAmount = parseAmount(
      (pricingDerived?.singerDiscountValue || 0)
      + (pricingDerived?.productionDiscountValue || 0)
    );

    const formatCurrency = (amount) => {
      if (amount === null || amount === undefined) return '';
      return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(amount);
    };

    const paymentLines = [];
    if (docType === 'invoice' || docType === 'quote') {
      if (current.balance_due_date) {
        paymentLines.push(`Balance due by ${formatDateDisplay(current.balance_due_date)}`);
      }
      if (depositAmount) {
        paymentLines.push(`Deposit: ${formatCurrency(depositAmount)}`);
      }
      if (balanceAmount && definition.invoice_variant === 'balance') {
        paymentLines.push(`Outstanding balance: ${formatCurrency(balanceAmount)}`);
      }
    }
    const paymentTerms = paymentLines.join('\n');

    const clientOverride = {
      name: current.client_name || '',
      email: current.client_email || '',
      phone: current.client_phone || '',
      address1: current.client_address1 || '',
      address2: current.client_address2 || '',
      address3: current.client_address3 || '',
      town: current.client_town || '',
      postcode: current.client_postcode || ''
    };

    const eventOverride = {
      type: current.event_type || '',
      event_name: current.event_type || '',
      event_date: current.event_date || '',
      startTime: current.event_start || '',
      endTime: current.event_end || '',
      venue_name: current.venue_name || '',
      venue_address1: current.venue_address1 || '',
      venue_address2: current.venue_address2 || '',
      venue_address3: current.venue_address3 || '',
      venue_town: current.venue_town || '',
      venue_postcode: current.venue_postcode || ''
    };

    const lineItems = [];
    const addLineItem = (label, amount, notes = '') => {
      const parsed = parseAmount(amount);
      if (parsed === null || parsed === 0) return;
      lineItems.push({
        date: current.event_date || '',
        description: label,
        notes,
        amount: parsed
      });
    };
    addLineItem('Performance fee', current.ahmen_fee);
    addLineItem('Production services', productionFees);
    addLineItem('Extras', extraFees);

    const payload = {
      business_id: numericBusinessId,
      doc_type: docType,
      definition_key: definition.key,
      document_date: new Date().toISOString(),
      total_amount: totalAmount ?? undefined,
      balance_amount: balanceAmount ?? undefined,
      balance_due: balanceAmount ?? undefined,
      balance_due_date: current.balance_due_date || undefined,
      balance_reminder_date: current.balance_reminder_date || undefined,
      deposit_amount: depositAmount ?? undefined,
      discount_amount: discountAmount ?? undefined,
      extra_fees: extraFees ?? undefined,
      production_fees: productionFees ?? undefined,
      service_types: current.service_types || '',
      specialist_singers: current.specialist_singers || '',
      notes: current.notes || '',
      payment_terms: paymentTerms,
      client_override: clientOverride,
      event_override: eventOverride,
      line_items: lineItems
    };

    try {
      payload.jobsheet_snapshot = JSON.parse(JSON.stringify(current));
    } catch (_err) {
      payload.jobsheet_snapshot = current;
    }

    try {
      payload.pricing_snapshot = pricingDerived ? JSON.parse(JSON.stringify(pricingDerived)) : null;
    } catch (_err) {
      payload.pricing_snapshot = pricingDerived || null;
    }

    // file_name_suffix removed
    if (definition.invoice_variant) payload.invoice_variant = definition.invoice_variant;

    if (docType === 'invoice') {
      payload.due_date = current.balance_due_date || current.event_date || undefined;
    }

    if (docType === 'quote') {
      payload.quote_meta = {
        validUntil: current.balance_due_date || '',
        includes: current.service_types || '',
        nextSteps: ''
      };
    }

    if (docType === 'contract') {
      payload.contract_meta = {
        terms: current.notes ? current.notes.split('\n').filter(Boolean) : [],
        signature: null
      };
    }

    if (!payload.footer && business?.document_footer) {
      payload.footer = business.document_footer;
    }

    return payload;
  }, [business, numericBusinessId, parseAmount, pricingDerived]);

  const validateDocumentRequest = useCallback((definition) => {
    const current = formStateRef.current || DEFAULT_JOBSHEET(numericBusinessId);
    const messages = [];

    if (!current.client_name?.trim()) messages.push('Add the client name.');
    if (!current.event_type?.trim()) messages.push('Add the event type.');
    if (!current.event_date) messages.push('Select the event date.');

    if (definition && definition.is_active === 0) {
      messages.push('This document type is inactive. Reactivate it before generating.');
    }

    return messages;
  }, [numericBusinessId]);

  const handlePopulateExcel = useCallback(async (requestedDefinitionKey) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setError('Document generation features are disabled.');
      return;
    }
    const previousSection = activeEditorSection || 'documents';
    if (jobsheetId != null && previousSection) {
      storeSection(jobsheetId, previousSection);
    }

    const targetKey = requestedDefinitionKey || selectedDefinitionKey;
    if (!targetKey) {
      setError('Select a document type to generate.');
      return;
    }

    const definition = findDefinitionByKey(targetKey);
    if (!definition) {
      setError('Document definition not found.');
      return;
    }

    const templatePath = definition.template_path || '';
    if (!templatePath) {
      setError('Choose the workbook template before generating.');
      openEditDefinitionModal(definition);
      return;
    }

    const errors = validateDocumentRequest(definition);
    if (errors.length) {
      setError(errors.join(' '));
      return;
    }

    const api = window.api;
    if (!api || typeof api.createDocument !== 'function') {
      setError('Unable to generate document: API unavailable');
      return;
    }

    const payload = buildDocumentPayload(definition);
    if (!payload) {
      setError('Unable to build document payload');
      return;
    }

    payload.template_path = templatePath;
    if (jobsheetId != null) {
      payload.jobsheet_id = Number(jobsheetId);
    }


    setDocumentGeneratingKey(definition.key);
    setError('');
    try {
      if (templatePath && templatePath.toLowerCase().endsWith('.xlsx')) {
        try {
          await window.api?.normalizeTemplate?.({ templatePath });
        } catch (normalizeErr) {
          console.warn('Failed to normalize template', normalizeErr);
        }
      }

      const result = await api.createDocument(payload);
      if (result?.file_path) {
        setLastOutputPath(result.file_path);
      }

      const suffix = result?.file_path ? ` saved to ${result.file_path}` : '';
      const baseLabel = definition.label || startCaseKey(definition.key);
      setMessage(`${baseLabel}${suffix}`.trim());

      if (Array.isArray(result?.additional_outputs) && result.additional_outputs.length) {
        const successes = result.additional_outputs.filter(item => item && item.success);
        if (successes.length) {
          const labels = successes.map(item => item.label || item.sheet || 'File').join(', ');
          setMessage(prev => `${prev ? `${prev}. ` : ''}Generated ${labels}.`);
          const firstPath = successes.find(item => item.file_path)?.file_path;
          if (firstPath) {
            setLastOutputPath(firstPath);
          }
        }
        const failures = result.additional_outputs.filter(item => !item?.success);
        if (failures.length) {
          const reasons = failures.map(item => `${item.sheet || 'Output'}: ${item.error || 'Unable to export'}`).join(' ');
          setError(reasons);
        }
      }

      window.api?.notifyJobsheetChange?.({
        type: 'documents-updated',
        businessId: numericBusinessId,
        jobsheetId: jobsheetId != null ? Number(jobsheetId) : null,
        document: result || null
      });
      setTimeout(() => setMessage(''), 4000);
      await refreshJobsheetDocuments();
      return result;
    } catch (err) {
      console.error('Failed to generate document', err);
      setError(err?.message || 'Unable to generate document');
      return null;
    } finally {
      setDocumentGeneratingKey(null);
      if (previousSection) {
        setActiveEditorSection(previousSection);
      }
    }
  }, [selectedDefinitionKey, findDefinitionByKey, validateDocumentRequest, buildDocumentPayload, jobsheetId, numericBusinessId, refreshJobsheetDocuments, setError, setMessage, openEditDefinitionModal, activeEditorSection, setActiveEditorSection, storeSection]);

  const handleRegenerateWorkbook = useCallback(async (definitionKey, existingDoc) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document generation features are disabled.');
      return;
    }

    const targetKey = definitionKey || 'workbook';
    const currentDoc = existingDoc
      || jobsheetDocuments.find(doc => doc?.definition_key === targetKey);

    const proceed = window.confirm('Regenerate the workbook using the latest jobsheet details?');
    if (!proceed) {
      return;
    }

    let removeFile = false;
    if (currentDoc?.file_path) {
      removeFile = window.confirm('Overwrite the existing workbook file on disk? Choose Cancel to keep the old file (a new copy will be created).');
    }

    try {
      setJobsheetDocumentsError('');
      if (currentDoc?.document_id != null && window.api?.deleteDocument) {
        await window.api.deleteDocument(currentDoc.document_id, { removeFile });
      } else if (removeFile && currentDoc?.file_path && window.api?.deleteDocumentByPath) {
        await window.api.deleteDocumentByPath({ businessId: numericBusinessId, absolutePath: currentDoc.file_path });
      }
    } catch (err) {
      console.error('Failed to remove existing workbook', err);
      setJobsheetDocumentsError(err?.message || 'Unable to remove existing workbook');
      return;
    }

    await refreshJobsheetDocuments();
    await handlePopulateExcel(targetKey);
  }, [jobsheetDocuments, numericBusinessId, refreshJobsheetDocuments, handlePopulateExcel]);


  const handleOpenOutputFolder = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setError('Document generation features are disabled.');
      return;
    }
    let folderPath = jobsheetDocumentsFolder || resolvedBusiness?.save_path;
    if (!folderPath) {
      try {
        const businessList = await window.api?.businessSettings?.();
        const match = (businessList || []).find(item => item.id === numericBusinessId);
        if (match?.save_path) {
          folderPath = match.save_path;
          setBusiness(prev => ({ ...(prev || {}), ...match }));
        }
      } catch (err) {
        console.error('Unable to reload business settings', err);
      }
    }
    if (!folderPath) {
      setError('Documents folder is not configured for this business.');
      return;
    }
    const response = await window.api?.openPath?.(folderPath);
    if (response && response.ok === false) {
      setError(response.message || 'Unable to open folder');
    }
  }, [jobsheetDocumentsFolder, resolvedBusiness, numericBusinessId, setBusiness]);

  const handleOpenOutputFile = useCallback(async () => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setError('Document generation features are disabled.');
      return;
    }
    if (!lastOutputPath) {
      setError('Generate the workbook before opening the file.');
      return;
    }
    const response = await window.api?.openPath?.(lastOutputPath);
    if (response && response.ok === false) {
      setError(response.message || 'Unable to open file');
    }
  }, [lastOutputPath]);

  useEffect(() => {
    if (loading) return () => {};
    if (jobsheetId) return () => {};
    const name = String(formState.client_name || '').trim();
    // Wait until user typed at least 2 chars and stopped typing briefly
    if (name.length < 2) { if (createDraftTimerRef.current) { clearTimeout(createDraftTimerRef.current); createDraftTimerRef.current = null; } return () => {}; }
    if (createDraftTimerRef.current) { clearTimeout(createDraftTimerRef.current); createDraftTimerRef.current = null; }
    createDraftTimerRef.current = setTimeout(async () => {
      if (creatingRef.current || jobsheetId) return;
      creatingRef.current = true;
      const api = window.api;
      if (!api || !api.addAhmenJobsheet) { creatingRef.current = false; return; }
      try {
        setSaving(true);
        const payload = preparePayload(formStateRef.current || formState, numericBusinessId);
        const newId = await api.addAhmenJobsheet(payload);
        setJobsheetId(newId);
        if (!isInline) {
          const url = new URL(window.location.href);
          url.searchParams.set('jobsheetId', newId);
          window.history.replaceState({}, '', url.toString());
        }
        window.api?.notifyJobsheetChange?.({
          type: 'jobsheet-created',
          businessId: numericBusinessId,
          jobsheetId: newId,
          snapshot: buildSnapshot({ ...(formStateRef.current || formState), jobsheet_id: newId }, newId)
        });
        setMessage('Draft created');
        setTimeout(() => setMessage(''), 1500);
        // Restore focus to client_name if user was typing
        setTimeout(() => {
          try {
            const el = document.querySelector('input[name="client_name"]');
            if (el) {
              const v = el.value || '';
              el.focus();
              const pos = v.length;
              if (el.setSelectionRange) el.setSelectionRange(pos, pos);
            }
          } catch (_) {}
        }, 50);
        initialLoadRef.current = true;
      } catch (err) {
        console.error('Failed to create jobsheet', err);
        setError(err?.message || 'Unable to create jobsheet');
      } finally {
        creatingRef.current = false;
        setSaving(false);
      }
    }, 650);
    return () => { if (createDraftTimerRef.current) { clearTimeout(createDraftTimerRef.current); createDraftTimerRef.current = null; } };
  }, [loading, jobsheetId, numericBusinessId, formState, isInline]);

  const handleSaveVenue = useCallback(async (overrideVenue) => {
    setVenueSaving(true);
    try {
      const source = overrideVenue ? {
        name: overrideVenue.name,
        address1: overrideVenue.address1,
        address2: overrideVenue.address2,
        address3: overrideVenue.address3,
        town: overrideVenue.town,
        postcode: overrideVenue.postcode,
        is_private: overrideVenue.is_private ? 1 : 0,
        venue_id: overrideVenue.venue_id || null
      } : {
        name: formState.venue_name,
        address1: formState.venue_address1,
        address2: formState.venue_address2,
        address3: formState.venue_address3,
        town: formState.venue_town,
        postcode: formState.venue_postcode,
        is_private: formState.venue_same_as_client ? 1 : 0,
        venue_id: formState.venue_id || null
      };

      if (!source.name?.trim()) {
        setError('Venue name is required to save.');
        return null;
      }

      const payload = {
        business_id: numericBusinessId,
        name: source.name,
        address1: source.address1,
        address2: source.address2,
        address3: source.address3,
        town: source.town,
        postcode: source.postcode,
        is_private: source.is_private,
        venue_id: source.venue_id || undefined
      };

      const api = window.api;
      if (!api || !api.saveAhmenVenue) {
        setError('Unable to save venue: API unavailable');
        return null;
      }

      const result = await api.saveAhmenVenue(payload);
      const savedVenueId = result?.venue_id ?? payload.venue_id ?? null;
      if (savedVenueId) {
        const optimisticVenue = normalizeVenues([
          {
            venue_id: savedVenueId,
            name: payload.name,
            address1: payload.address1,
            address2: payload.address2,
            address3: payload.address3,
            town: payload.town,
            postcode: payload.postcode,
            is_private: payload.is_private
          }
        ])[0];

        setVenues(prev => {
          const others = prev.filter(item => Number(item.venue_id) !== Number(savedVenueId));
          const nextList = [...others, optimisticVenue];
          nextList.sort((a, b) => a.name.localeCompare(b.name));
          return nextList;
        });

        const updatedVenues = await api.getAhmenVenues({ businessId: numericBusinessId });
        const normalized = normalizeVenues(updatedVenues);
        setVenues(normalized);

        const savedVenue = normalized.find(v => v.venue_id === savedVenueId);
        if (savedVenue) {
          setFormState(prev => applyDerivedFields({
            ...prev,
            venue_id: savedVenue.venue_id,
            venue_name: savedVenue.name,
            venue_address1: savedVenue.address1,
            venue_address2: savedVenue.address2,
            venue_address3: savedVenue.address3,
            venue_town: savedVenue.town,
            venue_postcode: savedVenue.postcode,
            venue_same_as_client: Boolean(savedVenue.is_private)
          }));
          setMessage('Venue saved');
          setTimeout(() => setMessage(''), 1500);
        }
      }

      window.api?.notifyJobsheetChange?.({
        type: 'jobsheet-updated',
        businessId: numericBusinessId,
        jobsheetId: jobsheetId || savedVenueId,
        snapshot: buildSnapshot(formState, jobsheetId || savedVenueId)
      });
     return savedVenueId;
    } catch (err) {
      console.error('Failed to save venue', err);
      setError(err?.message || 'Unable to save venue');
      return null;
    } finally {
      setVenueSaving(false);
    }
  }, [numericBusinessId, formState, jobsheetId]);

  const closeEditor = useCallback(() => {
    if (isInline) {
      onRequestClose?.();
    } else {
      window.close();
    }
  }, [isInline, onRequestClose]);

  const handleDelete = useCallback(async () => {
    if (!jobsheetId) {
      closeEditor();
      return;
    }
    const confirmed = window.confirm('Delete this jobsheet? This cannot be undone.');
    if (!confirmed) return;
    setSaving(true);
    try {
      const api = window.api;
      if (!api) {
        setError('Unable to delete jobsheet: API unavailable');
        setSaving(false);
        return;
      }
      // Prefer full cascade removal to avoid FK errors
      if (api.deleteJobsheetCompletely) {
        await api.deleteJobsheetCompletely({ businessId: numericBusinessId, jobsheetId, removeFiles: true });
      } else if (api.deleteAhmenJobsheet) {
        await api.deleteAhmenJobsheet(jobsheetId);
      } else {
        setError('Unable to delete jobsheet: API unavailable');
        setSaving(false);
        return;
      }
      window.api?.notifyJobsheetChange?.({ type: 'jobsheet-deleted', businessId: numericBusinessId, jobsheetId });
      closeEditor();
    } catch (err) {
      console.error('Failed to delete jobsheet', err);
      setError(err?.message || 'Unable to delete jobsheet');
    } finally {
      setSaving(false);
    }
  }, [jobsheetId, numericBusinessId, closeEditor]);

  useEffect(() => {
    const handler = () => {
      if (jobsheetId || formStateRef.current.client_name?.trim()) {
        window.api?.notifyJobsheetChange?.({
          type: 'jobsheet-updated',
          businessId: numericBusinessId,
          jobsheetId,
          snapshot: buildSnapshot(formStateRef.current, jobsheetId)
        });
      }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [numericBusinessId, jobsheetId, buildSnapshot]);

  const summarySingerFee = Number(formState.ahmen_fee) || (pricingDerived ? pricingDerived.singerNet : 0);
  const summaryProductionFee = Number(formState.production_fees) || (pricingDerived ? pricingDerived.productionNet : 0);
  const summaryTotal = pricingDerived ? pricingDerived.total : summarySingerFee + summaryProductionFee;
  const summaryCard = (
    <div className="bg-white border border-slate-200 rounded-lg px-5 py-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-5 text-sm text-slate-600">
      <div>
        <div className="text-xs uppercase tracking-wide text-slate-400">Client</div>
        <div className="text-base font-semibold text-slate-800">{formState.client_name || 'Untitled booking'}</div>
      </div>
      <div>
        <div className="text-xs uppercase tracking-wide text-slate-400">Event</div>
        <div className="text-base text-slate-700">{formState.event_type || '—'}</div>
        <div className="text-xs text-slate-500">{formatDateDisplay(formState.event_date)}</div>
      </div>
      <div>
        <div className="text-xs uppercase tracking-wide text-slate-400">Venue</div>
        <div className="text-base text-slate-700">{formState.venue_name || formState.venue_town || '—'}</div>
      </div>
      <div>
        <div className="text-xs uppercase tracking-wide text-slate-400">Fee</div>
        <div className="text-base font-semibold text-slate-800">{toCurrency(summaryTotal)}</div>
        <div className="text-xs text-slate-500">Singers {toCurrency(summarySingerFee)} · Production {toCurrency(summaryProductionFee)}</div>
      </div>
      <div>
        <div className="text-xs uppercase tracking-wide text-slate-400">Status</div>
        <span className={`inline-flex items-center rounded-full px-3 py-1 text-xs font-semibold ${STATUS_STYLES[formState.status] || STATUS_STYLES.enquiry}`}>
          {STATUS_OPTIONS.find(opt => opt.value === formState.status)?.label || 'Enquiry'}
        </span>
      </div>
    </div>
  );

  const workbookDefinition = useMemo(() => (
    (documentDefinitions || []).find(def => (def.doc_type || '').toLowerCase() === 'workbook') || null
  ), [documentDefinitions]);

  const editorContent = loading ? (
    <div className="bg-white rounded-lg border border-slate-200 p-6 text-center text-slate-500">Loading jobsheet…</div>
  ) : (
    <>
      {isInline ? (
        // Make the summary sticky in inline variant too
        <div id="jobsheet-sticky-header" className="sticky top-0 z-20 py-2 bg-slate-100/95 backdrop-blur">
          {summaryCard}
        </div>
      ) : (
        <div id="jobsheet-sticky-header" className="sticky top-0 z-20 -mx-6 px-6 pt-2 pb-4 bg-slate-100/95 backdrop-blur">
          {summaryCard}
        </div>
      )}
      <JobsheetEditor
        business={resolvedBusiness}
        businessId={numericBusinessId}
        jobsheetId={jobsheetId}
        formState={formState}
        onChange={setFormState}
        onDelete={handleDelete}
        saving={saving}
        deleting={false}
        hasExisting={Boolean(jobsheetId)}
        venues={venues}
        setVenues={setVenues}
        onSaveVenue={handleSaveVenue}
        venueSaving={venueSaving}
        setVenueSaving={setVenueSaving}
        pricingConfig={pricingConfig}
        pricingTotals={pricingDerived}
        onUpdateSingerPool={handleUpdateSingerPool}
        activeGroupKey={activeEditorSection}
        onActiveGroupChange={setActiveEditorSection}
        groups={fieldGroups}
        documents={jobsheetDocuments}
        documentsLoading={jobsheetDocumentsLoading}
        documentsError={jobsheetDocumentsError}
        documentDefinitions={documentDefinitions}
        definitionsLoading={documentDefinitionsLoading}
        onRefreshDocuments={refreshJobsheetDocuments}
        onGenerateDocument={handlePopulateExcel}
        onExportPdf={handleExportWorkbookPdf}
        onOpenDocumentFile={handleOpenDocumentFile}
        onRevealDocument={handleRevealDocument}
        onDeleteDocument={handleDeleteJobsheetDocument}
        documentFolder={jobsheetDocumentsFolder}
      />
      {/* Inline documents tab renders documents; legacy panel removed */}
    </>
  );

  const definitionModal = definitionModalOpen ? (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4 py-6">
      <div className="w-full max-w-2xl rounded-lg bg-white shadow-xl">
        <form
          onSubmit={event => {
            event.preventDefault();
            if (!definitionSaving) handleSaveDefinition();
          }}
          className="space-y-5 p-6"
        >
          <div className="flex items-start justify-between gap-4">
            <div>
              <h3 className="text-lg font-semibold text-slate-800">{definitionModalTitle}</h3>
              <p className="text-sm text-slate-500">Configure the template and behaviour for this document type.</p>
            </div>
            <button
              type="button"
              onClick={handleCloseDefinitionModal}
              className="text-slate-400 transition hover:text-slate-600"
              aria-label="Close"
            >
              ✕
            </button>
          </div>

          {definitionModalError ? (
            <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-600">
              {definitionModalError}
            </div>
          ) : null}

          <div className="grid gap-4 md:grid-cols-2">
            <label className="block text-sm font-medium text-slate-600">
              Label
              <input
                type="text"
                className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 disabled:cursor-not-allowed disabled:bg-slate-100"
                value={definitionDraft.label}
                onChange={event => handleDefinitionDraftChange('label', event.target.value)}
                placeholder="e.g. Statement of Work"
                disabled={modalIsLocked}
              />
            </label>
            <label className="block text-sm font-medium text-slate-600">
              Key
              <input
                type="text"
                className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 disabled:cursor-not-allowed disabled:bg-slate-100"
                value={definitionDraft.key}
                onChange={event => handleDefinitionDraftChange('key', event.target.value)}
                placeholder="e.g. statement_of_work"
                disabled={modalIsLocked}
              />
              <span className="mt-1 block text-xs text-slate-500">Lowercase letters, numbers, and underscores only.</span>
            </label>
            <label className="block text-sm font-medium text-slate-600">
              Document type
              <select
                className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 disabled:cursor-not-allowed disabled:bg-slate-100"
                value={definitionDraft.doc_type}
                onChange={event => handleDefinitionDraftChange('doc_type', event.target.value)}
                disabled={modalIsLocked}
              >
                {DOCUMENT_TYPE_OPTIONS.map(option => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
            <div className="space-y-3">
              {definitionDraft.doc_type === 'invoice' ? (
                <label className="block text-sm font-medium text-slate-600">
                  Invoice variant (optional)
                  <input
                    type="text"
                    className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                    value={definitionDraft.invoice_variant}
                    onChange={event => handleDefinitionDraftChange('invoice_variant', event.target.value)}
                    placeholder="e.g. deposit, balance"
                  />
                </label>
              ) : null}
              {/* Totals requirement removed; suffix removed */}
            </div>
            <label className="block text-sm font-medium text-slate-600 md:col-span-2">
              Description (optional)
              <textarea
                rows={3}
                className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                value={definitionDraft.description}
                onChange={event => handleDefinitionDraftChange('description', event.target.value)}
                placeholder="Explain what this template is used for."
              />
            </label>
          </div>

          <div className="space-y-3">
            <div className="rounded border border-slate-200 bg-slate-50 px-3 py-3 text-xs text-slate-600">
              <div className="font-medium text-slate-700">Template file</div>
              <p className="mt-1 break-all">
                {modalHasTemplate ? modalTemplatePath : 'No template selected yet.'}
              </p>
              <div className="mt-2 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handlePickDefinitionDraftTemplate}
                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  Choose file
                </button>
                <button
                  type="button"
                  onClick={handleOpenDraftTemplate}
                  disabled={!modalHasTemplate}
                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                >
                  Open file
                </button>
                <button
                  type="button"
                  onClick={handleClearDefinitionDraftTemplate}
                  disabled={!modalHasTemplate}
                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                >
                  Clear template
                </button>
                {modalSupportsNormalize ? (
                  <button
                    type="button"
                    onClick={handleNormalizeDraftTemplate}
                    className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                  >
                    Normalize
                  </button>
                ) : null}
              </div>
              {modalDocMeta.supportsNormalize ? (
                <p className="mt-2 text-[11px] text-slate-500">Excel templates are automatically normalized each time you generate a document.</p>
              ) : null}
            </div>

            <div className="flex flex-wrap gap-4 text-sm text-slate-600">
              <label className="inline-flex items-center gap-2">
                <input
                  type="checkbox"
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                  checked={Boolean(definitionDraft.is_active)}
                  onChange={event => handleDefinitionDraftChange('is_active', event.target.checked)}
                />
                Active
              </label>
            </div>

            {definitionModalMode === 'edit' && modalIsLocked ? (
              <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-[11px] text-slate-500">
                This definition is part of the default set and cannot be deleted. You can still attach a different template.
              </div>
            ) : null}
          </div>

          <div className="flex items-center justify-between gap-2">
            {definitionModalMode === 'edit' && !modalIsLocked ? (
              <button
                type="button"
                onClick={() => handleDeleteDefinition(definitionDraft)}
                className="inline-flex items-center rounded border border-red-200 px-4 py-2 text-sm font-medium text-red-600 hover:bg-red-50"
              >
                Delete
              </button>
            ) : <span />}
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleCloseDefinitionModal}
                className="inline-flex items-center rounded border border-slate-300 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="submit"
                disabled={definitionSaving}
                className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500 disabled:cursor-not-allowed disabled:opacity-60"
              >
                {definitionSaving ? 'Saving…' : 'Save definition'}
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  ) : null;

  const editorToasts = [];
  if (error) editorToasts.push({ id: 'jobsheet-error', tone: 'error', text: error });
  if (message) editorToasts.push({ id: 'jobsheet-message', tone: 'success', text: message });

  if (isInline) {
    const inlineStatus = saving ? 'Saving…' : '';
    const inlineMessageVisible = !error && Boolean(inlineStatus);
    const inlineDisplay = inlineStatus || '\u00A0';
    return (
      <div className="space-y-4 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 xl:px-10 py-4 sm:py-6">
        <ToastOverlay notices={editorToasts} />
        <div className="min-h-[2.5rem]" aria-live="polite" aria-atomic="true">
          <div
            className={`rounded border border-slate-200 bg-slate-50 px-4 py-2 text-xs font-medium text-slate-600 transition duration-200 ${inlineMessageVisible ? 'opacity-100 translate-y-0' : 'opacity-0 -translate-y-1 pointer-events-none'}`}
          >
            {inlineDisplay}
          </div>
        </div>
        {editorContent}
        {definitionModal}
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100">
      <ToastOverlay notices={editorToasts} />
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold text-slate-800">{resolvedBusiness.business_name}</h1>
            <p className="text-sm text-slate-500">Jobsheet editor · changes save automatically.</p>
          </div>
          <div className="text-sm text-slate-500">
            {saving ? 'Saving…' : message || ''}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6 space-y-4">
        {editorContent}
      </main>

      {definitionModal}
    </div>
  );
}

function mapApiToForm(apiData, businessId) {
  const base = DEFAULT_JOBSHEET(businessId);
  Object.entries(apiData || {}).forEach(([key, value]) => {
    if (key === 'pricing_selected_singers') {
      base[key] = parseSelectedSingers(value);
      return;
    }
    if (key === 'pricing_production_items') {
      base[key] = parseProductionItems(value);
      return;
    }
    if (key === 'venue_same_as_client') {
      base[key] = Boolean(value);
      return;
    }
    if (key === 'status') {
      const normalized = normalizeStatus(value) || 'enquiry';
      base[key] = normalized;
      return;
    }
    if (key === 'pricing_discount_type') {
      base[key] = value || 'amount';
      return;
    }
    if (key === 'pricing_discount_value' || key === 'pricing_production_total') {
      base[key] = value != null ? String(value) : '';
      return;
    }
    if (key === 'pricing_production_subtotal' || key === 'pricing_production_discount_value') {
      base[key] = value != null ? String(value) : '';
      return;
    }
    if (key === 'pricing_production_discount' || key === 'pricing_production_discount_type') {
      base[key] = value != null ? String(value) : '';
      return;
    }
    base[key] = value ?? base[key] ?? '';
  });
  if (!base.pricing_discount_type) base.pricing_discount_type = 'amount';
  if (base.pricing_discount === undefined || base.pricing_discount === null) base.pricing_discount = '';
  if (!base.pricing_discount_value) base.pricing_discount_value = '';
  if (!base.pricing_production_discount_type) base.pricing_production_discount_type = 'amount';
  if (base.pricing_production_discount === undefined || base.pricing_production_discount === null) base.pricing_production_discount = '';
  if (!base.pricing_production_discount_value) base.pricing_production_discount_value = '';
  if (!base.pricing_production_subtotal) base.pricing_production_subtotal = '';
  if (!base.pricing_production_total && base.production_fees != null) {
    base.pricing_production_total = String(base.production_fees);
  }
  if (!base.pricing_production_subtotal && base.production_fees != null) {
    base.pricing_production_subtotal = String(base.production_fees);
  }
  return applyDerivedFields(base);
}

function App() {
  const searchParams = useMemo(() => new URLSearchParams(window.location.search), []);
  const mode = searchParams.get('mode');

  if (mode === 'jobsheet') {
    const businessIdParam = searchParams.get('businessId');
    const businessNameParam = searchParams.get('businessName') || '';
    const jobsheetIdParam = searchParams.get('jobsheetId');
    return (
      <JobsheetEditorWindow
        businessId={businessIdParam}
        businessName={businessNameParam}
        initialJobsheetId={jobsheetIdParam}
      />
    );
  }

  const [businesses, setBusinesses] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [selectedBusiness, setSelectedBusiness] = useState(null);

  useEffect(() => {
    let mounted = true;
    const load = async () => {
      setLoading(true);
      setError('');
      try {
        const data = await window.api.businessSettings();
        if (!mounted) return;
        const businessList = Array.isArray(data) ? data : [];
        // Hide MCMS (id=1); prefer AhMen (id=2)
        const filtered = businessList.filter(biz => biz && biz.id !== 1);
        setBusinesses(filtered);
        // Auto-select AhMen when present
        const preferred = filtered.find(biz => biz.id === 2) || filtered[0] || null;
        if (preferred) {
          storeLastBusinessId(preferred.id);
          setSelectedBusiness(preferred);
        }
      } catch (err) {
        if (!mounted) return;
        console.error('Failed to load businesses', err);
        setError(err?.message || 'Unable to load businesses');
      } finally {
        if (mounted) setLoading(false);
      }
    };
    load();
    return () => {
      mounted = false;
    };
  }, []);

  const handleSelectBusiness = (business) => {
    if (!business) return;
    storeLastBusinessId(business.id);
    setSelectedBusiness(business);
  };

  const handleBusinessUpdated = useCallback((updatedBusiness) => {
    if (!updatedBusiness) return;
    setBusinesses(prev => prev.map(biz => (biz.id === updatedBusiness.id ? { ...biz, ...updatedBusiness } : biz)));
    setSelectedBusiness(updatedBusiness);
  }, [setBusinesses, setSelectedBusiness]);

  if (!selectedBusiness) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-8">
        <div className="text-slate-600 text-sm">Loading…</div>
      </div>
    );
  }

  return (
    <BusinessWorkspace
      business={selectedBusiness}
      onBusinessUpdate={handleBusinessUpdated}
    />
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = createRoot(rootElement);
  root.render(<App />);
}

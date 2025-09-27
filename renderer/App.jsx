import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import TemplatesManager from './components/TemplatesManager';
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

const DOCUMENT_GROUP_OPTIONS = [
  { value: 'none', label: 'All Documents' },
  { value: 'doc_type', label: 'Document Type' },
  { value: 'client', label: 'Client' },
  { value: 'event_date', label: 'Event Date' }
];

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
  templates: '📁',
  settings: '⚙️'
};

const WORKSPACE_SECTIONS = [
  { key: 'jobsheets', label: 'Jobsheets', description: 'Bookings and statuses', icon: WORKSPACE_ICON_MAP.jobsheets },
  { key: 'documents', label: 'Documents', description: 'Browse and manage files', icon: WORKSPACE_ICON_MAP.documents },
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
    file_suffix: '',
    invoice_variant: '',
    template_path: '',
    requires_total: 1,
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
  confirmed: 'bg-green-100 text-green-800 border border-green-200',
  completed: 'bg-gray-200 text-gray-700 border border-gray-300'
};

const STATUS_ROW_CLASSES = {
  enquiry: 'bg-yellow-100',
  quoted: 'bg-blue-100',
  confirmed: 'bg-green-100',
  completed: 'bg-gray-200'
};

const ACTIVE_STATUS_ROW_CLASSES = {
  enquiry: 'bg-yellow-400',
  quoted: 'bg-blue-400',
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
  }
};

const GROUP_ICON_MAP = {
  client: '👤',
  event: '🎉',
  venue: '📍',
  pricing: '🎶',
  production: '🎛️',
  billing: '💷',
  services: '📝'
};

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

function IconButton({ label, onClick, disabled, className = '', children }) {
  const handleClick = useCallback((event) => {
    event.stopPropagation();
    onClick?.(event);
  }, [onClick]);

  return (
    <button
      type="button"
      onClick={handleClick}
      disabled={disabled}
      className={`inline-flex h-8 w-8 items-center justify-center rounded border border-slate-300 text-slate-600 transition hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-60 ${className}`}
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
  onCollapsedChange
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
    setCollapsedNodes(new Set());
    setSelectedNodeId(null);
  }, [safeRootPath, root]);

  const toggleFolder = useCallback((nodeId) => {
    setCollapsedNodes(prev => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      return next;
    });
  }, []);

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
        onClick={() => setSelectedNodeId(nodeId)}
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

function BusinessChooser({ businesses, loading, error, onSelect }) {
  return (
    <div className="min-h-screen bg-slate-100 flex flex-col items-center justify-center p-8">
      <div className="max-w-4xl w-full">
        <h1 className="text-3xl font-bold text-slate-800 text-center mb-8">Choose a business to continue</h1>
        {error ? (
          <div className="bg-red-100 text-red-700 p-4 rounded mb-6">{error}</div>
        ) : null}
        {loading ? (
          <div className="text-center text-slate-600">Loading businesses…</div>
        ) : (
          <div className="grid gap-6 md:grid-cols-2">
            {businesses.map(biz => (
              <button
                key={biz.id}
                onClick={() => onSelect(biz)}
                className="rounded-xl bg-white shadow-md hover:shadow-lg transition-shadow text-left p-6 border border-slate-200"
              >
                <div className="text-sm uppercase tracking-wide text-slate-500 mb-2">Business</div>
                <div className="text-2xl font-semibold text-slate-800 mb-4">{biz.business_name}</div>
                <dl className="text-sm text-slate-600 space-y-2">
                  <div>
                    <dt className="font-medium text-slate-500">Invoices to date</dt>
                    <dd>{biz.last_invoice_number ?? '—'}</dd>
                  </div>
                  <div>
                    <dt className="font-medium text-slate-500">Quotes to date</dt>
                    <dd>{biz.last_quote_number ?? '—'}</dd>
                  </div>
                  <div>
                    <dt className="font-medium text-slate-500">Documents folder</dt>
                    <dd className="truncate" title={biz.save_path}>{biz.save_path || 'Not configured'}</dd>
                  </div>
                </dl>
              </button>
            ))}
            {!businesses.length ? (
              <div className="col-span-full text-center text-slate-500">No businesses found. Populate the database first.</div>
            ) : null}
          </div>
        )}
      </div>
    </div>
  );
}

function JobsheetList({
  jobsheets,
  onOpen,
  onNew,
  onDelete,
  onStatusChange,
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
        className="cursor-pointer"
      >
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} ${firstCellExtras}`}>
          <div className="flex items-center gap-3">
            {isActive ? <span className="h-8 w-1 rounded-full bg-indigo-600" /> : null}
            <span className="font-medium text-slate-800 whitespace-nowrap">{sheet.client_name || 'Untitled booking'}</span>
          </div>
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder}`}>
          {sheet.event_type || '—'}
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} whitespace-nowrap`}>
          {formatDateDisplay(sheet.event_date)}
        </td>
        <td className={`${rowBackground} ${baseCellClass} ${verticalBorder} truncate`}>
          {sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '—'}
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
          <div className="flex justify-end">
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
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h2 className="text-lg font-semibold text-slate-700">Jobsheets</h2>
            <p className="text-sm text-slate-500">{summaryLabel}</p>
          </div>
          <button
            onClick={onNew}
            className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-medium px-3 py-2 rounded"
          >
            + New Jobsheet
          </button>
        </div>
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
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
          </div>
        </div>
      </div>
      <div className="flex-1 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {loading ? (
          <div className="p-6 text-center text-slate-500">Loading…</div>
        ) : !sortedJobsheets.length ? (
          <div className="p-6 text-center text-slate-500">{hasActiveFilters ? 'No jobsheets match your filters yet. Adjust the search or status filter to see more results.' : 'No jobsheets yet. Create your first one!'}</div>
        ) : (
          <div className="overflow-y-auto">
            <table className="min-w-full text-sm border-separate border-spacing-y-2">
              <thead>
                <tr className="bg-slate-50">
                  {JOBSHEET_COLUMNS.map(column => {
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
                {sortedJobsheets.map(sheet => renderDataRow(sheet))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
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
      <div className="rounded-lg border border-slate-200 bg-slate-100 shadow-sm overflow-hidden">
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
  const selectedService = serviceTypes.find(type => String(type.id) === selectedServiceId) || null;
  const lastServiceIdRef = useRef('');

  useEffect(() => {
    const currentServiceId = selectedService ? String(selectedService.id) : '';
    if (!currentServiceId) {
      if (selectedEntries.length) updateSelected([]);
      lastServiceIdRef.current = '';
      return;
    }

    if (currentServiceId !== lastServiceIdRef.current) {
      lastServiceIdRef.current = currentServiceId;
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
  }, [selectedService, sortedSingers, selectedEntries, poolMap, updateSelected]);

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
                  className={`flex flex-wrap items-center gap-3 rounded border px-3 py-2 text-sm transition ${
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

                  <div className="flex min-w-[10rem] flex-1 items-center">
                    <span className={`font-medium leading-tight ${isSelected ? 'text-white' : 'text-slate-700'}`}>
                      {singer.name || 'Unnamed singer'}
                    </span>
                  </div>

                  <label
                    className={`flex flex-shrink-0 items-center gap-1 text-xs uppercase tracking-wide ${
                      isSelected ? 'text-white/80' : 'text-slate-500'
                    }`}
                  >
                    <span>Fee</span>
                    <div className="relative flex items-center">
                      <span className={`pointer-events-none absolute left-2 ${isSelected ? 'text-white/70' : 'text-slate-400'}`}>£</span>
                      <input
                        type="number"
                        step="0.01"
                        className={`w-28 rounded border px-5 py-1 text-sm focus:outline-none focus:ring-2 ${
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

                  <div className="ml-auto flex items-center gap-2">
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

function Field({ label, type = 'text', value, onChange, readOnly, hint, rows = 3, step, component, options }) {
  const common = {
    className: 'mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500',
    value: value ?? '',
    onChange: (event) => onChange(event.target.value),
    readOnly,
    disabled: readOnly,
    step
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
  groups
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

  const [savedVenueId, setSavedVenueId] = useState(() => (
    formState.venue_id ? String(formState.venue_id) : ''
  ));

  useEffect(() => {
    setSavedVenueId(formState.venue_id ? String(formState.venue_id) : '');
  }, [formState.venue_id]);

  const [showVenueModal, setShowVenueModal] = useState(false);
  const [venueDraft, setVenueDraft] = useState(() => buildVenueDraft());

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

  const [activeGroupKey, setActiveGroupKey] = useState(() => {
    if (activeGroupKeyProp) return activeGroupKeyProp;
    const defaultGroup = resolvedGroups.find(group => group.defaultOpen) || resolvedGroups[0];
    return defaultGroup ? defaultGroup.key : null;
  });

  useEffect(() => {
    if (activeGroupKeyProp) {
      setActiveGroupKey(activeGroupKeyProp);
      return;
    }
    setActiveGroupKey(prev => {
      if (prev && resolvedGroups.some(group => group.key === prev)) return prev;
      const fallbackGroup = resolvedGroups.find(group => group.defaultOpen) || resolvedGroups[0] || null;
      return fallbackGroup ? fallbackGroup.key : null;
    });
  }, [resolvedGroups, activeGroupKeyProp]);

  const setGroupKey = useCallback((nextKey) => {
    if (!nextKey) return;
    if (!resolvedGroups.some(group => group.key === nextKey)) return;
    setActiveGroupKey(nextKey);
    onActiveGroupChange?.(nextKey);
  }, [resolvedGroups, onActiveGroupChange]);

  const activeGroup = useMemo(() => (
    resolvedGroups.find(group => group.key === activeGroupKey) || null
  ), [resolvedGroups, activeGroupKey]);

  useEffect(() => {
    if (!activeGroup && resolvedGroups.length) {
      const fallbackGroup = resolvedGroups.find(group => group.defaultOpen) || resolvedGroups[0] || null;
      if (fallbackGroup) setGroupKey(fallbackGroup.key);
    }
  }, [activeGroup, resolvedGroups, setGroupKey]);

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
        <nav className="lg:w-64 flex-shrink-0 lg:sticky lg:top-4 self-start">
          <div className="space-y-2" role="tablist" aria-orientation="vertical">
            {resolvedGroups.map(group => {
              const isActive = activeGroup?.key === group.key;
              const icon = group.icon ?? getGroupIcon(group.key);
              return (
                <button
                  key={group.key}
                  type="button"
                  role="tab"
                  aria-selected={isActive}
                  onClick={() => setGroupKey(group.key)}
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
          {activeGroup ? (
            <section className="bg-white border border-slate-200 rounded-lg p-5 space-y-5">
              <div>
                <h3 className="text-lg font-semibold text-slate-700">{activeGroup.title}</h3>
                {activeGroup.description ? (
                  <p className="mt-1 text-sm text-slate-500">{activeGroup.description}</p>
                ) : null}
              </div>
              <div className="space-y-4">
                {activeGroup.fields.map(field => {
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
                      type={field.type || 'text'}
                      step={field.step}
                      rows={field.rows}
                      hint={field.hint}
                      readOnly={field.name === 'venue_name' ? Boolean(formState.venue_same_as_client) : field.readOnly}
                      component={field.component}
                      options={field.options}
                      value={resolvedValue}
                      onChange={value => handleFieldChange(
                        field.name,
                        field.type === 'checkbox' ? Boolean(value) : value
                      )}
                    />
                  );
                })}
              </div>
            </section>
          ) : (
            <div className="rounded-lg border border-slate-200 bg-white p-5 text-sm text-slate-500">
              No sections available.
            </div>
          )}
        </div>
      </div>

      <div className="flex items-center justify-end text-sm text-slate-500 min-h-[1.5rem]">
        {saving ? 'Saving changes…' : null}
      </div>
      </div>

      {showVenueModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl">
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
                    window.api?.openExternal?.(url) || window.open(url, '_blank');
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
                    window.api?.openExternal?.(url) || window.open(url, '_blank');
                  }}
                >
                  Search Maps
                </button>
              </div>
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

function BusinessWorkspace({ business, onSwitch, onBusinessUpdate }) {
  const [jobsheets, setJobsheets] = useState([]);
  const [listLoading, setListLoading] = useState(true);
  const [sortConfig, setSortConfig] = useState({ key: 'event_date', direction: 'desc' });
  const [deletingId, setDeletingId] = useState(null);
  const [statusUpdatingId, setStatusUpdatingId] = useState(null);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [activeJobsheetId, setActiveJobsheetId] = useState(null);
  const [inlineEditorVisible, setInlineEditorVisible] = useState(false);
  const [inlineEditorTargetId, setInlineEditorTargetId] = useState(null);
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

  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      window.localStorage.setItem(DOCUMENT_TREE_COLLAPSE_KEY, documentTreeCollapsed ? 'true' : 'false');
    } catch (err) {
      console.warn('Unable to store document tree collapse preference', err);
    }
  }, [documentTreeCollapsed]);

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
      const data = await api.getAhmenJobsheets({ businessId: business.id });
      const mapped = (data || []).map(normalizeJobsheet);
      setJobsheets(mapped);

      const currentActive = activeJobsheetIdRef.current;
      if (currentActive != null) {
        const exists = mapped.some(job => job?.jobsheet_id != null && Number(job.jobsheet_id) === currentActive);
        if (exists) {
          setActiveJobsheetId(currentActive);
          setInlineEditorTargetId(currentActive);
          setInlineEditorVisible(true);
        }
      }
    } catch (err) {
      console.error('Failed to refresh jobsheets', err);
      setError(err?.message || 'Unable to refresh jobsheets');
    } finally {
      setListLoading(false);
    }
  }, [business.id, normalizeJobsheet]);

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
        if (!api || typeof api.getDocuments !== 'function') {
          throw new Error('Unable to load documents: API unavailable');
        }
        const data = await api.getDocuments({ businessId: business.id });
        setDocuments(Array.isArray(data) ? data : []);
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
        fileAvailable: Boolean(doc.file_path)
      };
    });
  }, [documents]);

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
        doc.displayClient,
        doc.displayEvent,
        doc.statusLabel,
        doc.formattedEventDate,
        doc.formattedDocumentDate,
        doc.createdAtDisplay,
        doc.createdAtFull,
        doc.doc_type,
        doc.file_path,
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
              const docTitle = doc.typeLabel + (doc.number ? ` #${doc.number}` : '');
              const fileName = doc.file_path ? doc.file_path.split(/[\\/]+/).filter(Boolean).pop() : '';
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
                      const tooltipText = doc.file_path || docTitle;

                      cell = (
                        <div className="flex items-center gap-3">
                          <span className="text-lg" role="img" aria-label={doc.typeLabel}>{getDocumentIcon(doc.doc_type)}</span>
                          <div
                            className="text-sm font-medium text-slate-700 truncate"
                            title={tooltipText}
                          >
                            {fileName || docTitle}
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
                      cell = (
                        <div className="flex flex-wrap justify-end gap-1.5">
                          <IconButton
                            label="Open document"
                            onClick={() => handleOpenDocumentFile(doc.file_path)}
                            disabled={!doc.fileAvailable}
                          >
                            <OpenIcon />
                          </IconButton>
                          <IconButton
                            label="Reveal document in Finder"
                            onClick={() => handleRevealDocument(doc.file_path)}
                            disabled={!doc.fileAvailable}
                          >
                            <RevealIcon />
                          </IconButton>
                          <IconButton
                            label="Delete document"
                            onClick={() => handleDeleteDocumentRecord(doc)}
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

  const documentsContent = useMemo(() => {
    if (!filteredDocuments.length) {
      return <div className="rounded border border-slate-200 bg-slate-50 px-4 py-6 text-center text-sm text-slate-500">{emptyStateMessage}</div>;
    }

    if (documentsGroup === 'none') {
      return renderDocumentTable(filteredDocuments);
    }

    return groupedDocuments.map(group => (
      <div key={group.key || 'group'} className="space-y-2">
        <h3 className="text-sm font-semibold text-slate-600">{group.label || 'Other'}</h3>
        {renderDocumentTable(group.items)}
      </div>
    ));
  }, [documentsGroup, filteredDocuments, groupedDocuments, renderDocumentTable, emptyStateMessage]);

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

  const handleNew = useCallback(() => {
    setActiveJobsheetId(null);
    setInlineEditorTargetId(null);
    setInlineEditorVisible(true);
    setInlineEditorSession(prev => prev + 1);
  }, []);

  const handleOpenExisting = useCallback((jobsheetId) => {
    if (!jobsheetId) return;
    const numericId = Number(jobsheetId);
    setActiveJobsheetId(numericId);
    setInlineEditorTargetId(numericId);
    setInlineEditorVisible(true);
    setInlineEditorSession(prev => (numericId !== inlineEditorTargetId ? prev + 1 : prev));
  }, [inlineEditorTargetId]);

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
      await api.deleteAhmenJobsheet(jobsheetId);
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

  return (
    <div className="min-h-screen bg-slate-100">
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold text-slate-800">{business.business_name}</h1>
            <p className="text-sm text-slate-500">Manage jobsheets, documents, and templates in one workspace.</p>
          </div>
          <button
            onClick={onSwitch}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
          >
            Switch business
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        {error ? <div className="rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
        {message ? <div className="rounded border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-700">{message}</div> : null}

        <div className="flex flex-col gap-6 lg:flex-row">
          <nav className="sticky top-4 z-30 flex-shrink-0 self-start lg:w-64">
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
                  jobsheets={jobsheets}
                  onOpen={handleOpenExisting}
                  onNew={handleNew}
                  onDelete={handleDelete}
                  onStatusChange={handleStatusChange}
                  loading={listLoading}
                  deletingId={deletingId}
                  statusUpdatingId={statusUpdatingId}
                  sortConfig={sortConfig}
                  onSort={handleSort}
                  activeJobsheetId={activeJobsheetId}
                />
                <InlineJobsheetEditorPanel
                  business={business}
                  visible={inlineEditorVisible}
                  jobsheetId={inlineEditorTargetId}
                  sessionKey={inlineEditorKey}
                  onClose={handleCloseInlineEditor}
                  onOpenInWindow={handlePopoutEditor}
                />
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
                  />
                  <div className="space-y-4">
                    <div className="rounded-lg border border-slate-200 bg-white p-4 space-y-4">
                      <div className="flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
                        <div>
                          <h2 className="text-lg font-semibold text-slate-700">Documents</h2>
                          <p className="text-sm text-slate-500">
                            {headerSubtitle}
                            {documentsLoading ? <span className="ml-2 text-xs text-slate-400">Loading…</span> : null}
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

            {workspaceSection === 'settings' ? (
              <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-700">Business settings</h2>
                  <p className="text-sm text-slate-500">Update folders and review business information.</p>
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

                <div className="grid gap-3 md:grid-cols-2">
                  <div className="rounded border border-slate-200 p-4">
                    <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Last invoice number</div>
                    <div className="mt-1 text-sm text-slate-700">{business.last_invoice_number ?? '—'}</div>
                  </div>
                  <div className="rounded border border-slate-200 p-4">
                    <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Last quote number</div>
                    <div className="mt-1 text-sm text-slate-700">{business.last_quote_number ?? '—'}</div>
                  </div>
                </div>

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

function JobsheetDocumentsPanel({
  jobsheetId,
  documents,
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
  lastOutputPath
}) {
  if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) return null;

  const list = Array.isArray(documents) ? documents : [];
  const workbookKey = workbookDefinition?.key || 'workbook';
  const hasTemplate = Boolean(workbookDefinition?.template_path);
  const isGenerating = generatingKey === workbookKey;
  const disableGenerate = !jobsheetId || !hasTemplate || isGenerating || definitionsLoading;
  const existingWorkbook = list.find(doc => (doc.definition_key || 'workbook') === workbookKey || (doc.doc_type === 'workbook' && !doc.definition_key));

  const documentLabel = (doc) => {
    return doc?.definition_label
      || (doc?.definition_key ? startCaseKey(doc.definition_key) : 'Document');
  };

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
        <div>
          <h2 className="text-lg font-semibold text-slate-700">Documents</h2>
          <p className="text-sm text-slate-500">Generate the master workbook for this jobsheet and open the latest exports.</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={onRefresh}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            Refresh list
          </button>
          <button
            type="button"
            onClick={onOpenOutputFolder}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            Open documents folder
          </button>
          <button
            type="button"
            onClick={onOpenOutputFile}
            disabled={!lastOutputPath}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
          >
            Open latest export
          </button>
        </div>
      </div>

      {error ? (
        <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-600">{error}</div>
      ) : null}

      <div className="rounded border border-slate-200 bg-slate-50 px-3 py-3 text-sm text-slate-600">
        {definitionsLoading ? (
          <div>Loading template details…</div>
        ) : workbookDefinition ? (
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <div className="font-medium text-slate-700">Workbook template</div>
              <p className="text-xs text-slate-500 break-all">
                {workbookDefinition.template_path || 'No template selected yet. Use the Templates tab to choose a file.'}
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => onEditTemplate?.(workbookDefinition)}
                className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
              >
                Edit template settings
              </button>
              <button
                type="button"
                disabled={!hasTemplate}
                onClick={() => onOpenTemplate?.(workbookDefinition)}
                className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
              >
                Open template file
              </button>
              {existingWorkbook ? (
                <button
                  type="button"
                  onClick={() => onRegenerate?.(workbookKey, existingWorkbook)}
                  disabled={disableGenerate}
                  className="inline-flex items-center rounded border border-indigo-200 px-3 py-1.5 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-60"
                >
                  {isGenerating ? 'Regenerating…' : 'Regenerate workbook'}
                </button>
              ) : null}
              <button
                type="button"
                onClick={() => onGenerate(workbookKey)}
                disabled={disableGenerate}
                className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-xs font-semibold text-white hover:bg-indigo-500 disabled:cursor-not-allowed disabled:opacity-60"
              >
                {isGenerating ? 'Generating…' : 'Generate workbook'}
              </button>
            </div>
          </div>
        ) : (
          <div>
            Configure the workbook template in the Templates tab before generating documents.
          </div>
        )}
        {!jobsheetId ? (
          <p className="mt-2 text-xs text-slate-500">Save the jobsheet before generating the workbook.</p>
        ) : null}
      </div>

      <div className="space-y-3">
        <div className="flex items-center justify-between text-sm">
          <div className="font-medium text-slate-700">Generated files</div>
          {loading ? <span className="text-xs text-slate-500">Loading…</span> : null}
        </div>
        {list.length === 0 && !loading ? (
          <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-500">
            No documents generated yet.
          </div>
        ) : null}
        {list.length > 0 ? (
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
                {list.map(doc => {
                  const title = documentLabel(doc);
                  const fileName = doc.file_path ? doc.file_path.split(/[\\/]+/).filter(Boolean).pop() : '';
                  const createdDisplay = doc?.created_at ? formatTimestampDisplay(doc.created_at) : '—';
                  const tooltip = doc.file_path || title;
                  return (
                    <tr key={doc.document_id}>
                      <td className="px-3 py-2 align-top">
                        <div className="flex items-center gap-3">
                          <span className="text-lg" role="img" aria-label={title}>{getDocumentIcon(doc.doc_type)}</span>
                          <div className="text-xs font-medium text-slate-700 truncate" style={{ maxWidth: '24rem' }} title={tooltip}>
                            {fileName || title}
                          </div>
                        </div>
                      </td>
                      <td className="px-3 py-2 align-top text-sm text-slate-600">{createdDisplay}</td>
                      <td className="px-3 py-2 align-top">
                        <div className="flex flex-wrap justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => onOpenFile?.(doc.file_path)}
                            disabled={!doc.file_path}
                            className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                          >
                            Open
                          </button>
                          <button
                            type="button"
                            onClick={() => onRevealFile?.(doc.file_path)}
                            disabled={!doc.file_path}
                            className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                          >
                            Reveal in Finder
                          </button>
                          <button
                            type="button"
                            onClick={() => onDelete?.(doc)}
                            className="inline-flex items-center rounded border border-rose-200 px-2 py-1 text-xs font-medium text-rose-600 hover:bg-rose-50"
                          >
                            Delete
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
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

  const autoSaveTimer = useRef(null);
  const initialLoadRef = useRef(true);
  const creatingRef = useRef(initialResolvedJobsheetId == null);
  const previousJobsheetIdRef = useRef(initialResolvedJobsheetId);

  const storagePrefix = useMemo(() => (
    `jobsheetEditor:${isInline ? 'inline' : 'window'}:${numericBusinessId}:`
  ), [isInline, numericBusinessId]);

  const getStoredSection = useCallback((jobsheetValue) => {
    const id = jobsheetValue != null ? Number(jobsheetValue) : null;
    if (!id || Number.isNaN(id)) return null;
    try {
      return window.sessionStorage.getItem(`${storagePrefix}${id}`) || null;
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
    storeSection(id, activeEditorSection);
  }, [jobsheetId, activeEditorSection, storeSection]);

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
        const primary = list.find(def => def.is_primary);
        const fallback = primary || list[0];
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
      setJobsheetDocumentsLoading(false);
      setJobsheetDocumentsError('');
      return;
    }
    if (!jobsheetId) {
      setJobsheetDocuments([]);
      setJobsheetDocumentsLoading(false);
      setJobsheetDocumentsError('');
      return;
    }
    setJobsheetDocumentsLoading(true);
    setJobsheetDocumentsError('');
    try {
      const api = window.api;
      if (!api || typeof api.getDocuments !== 'function') {
        throw new Error('Unable to load documents: API unavailable');
      }
      const fetchDocuments = async () => {
        const data = await api.getDocuments({ businessId: numericBusinessId });
        return Array.isArray(data) ? data : [];
      };

      const normalizedJobsheetId = jobsheetId != null ? Number(jobsheetId) : null;
      const currentState = formStateRef.current || DEFAULT_JOBSHEET(numericBusinessId);

      const filterForJobsheet = (docs) => {
        return docs.filter(doc => {
          const docJobsheetId = doc?.jobsheet_id != null ? Number(doc.jobsheet_id) : null;
          if (normalizedJobsheetId != null && docJobsheetId === normalizedJobsheetId) {
            return true;
          }
          if (normalizedJobsheetId != null && docJobsheetId != null && docJobsheetId !== normalizedJobsheetId) {
            return false;
          }
          if (docJobsheetId == null && normalizedJobsheetId != null) {
            return matchesDocumentToJobsheet(doc, currentState);
          }
          return false;
        });
      };

      const initialDocs = await fetchDocuments();
      let filtered = filterForJobsheet(initialDocs);

      if (normalizedJobsheetId != null && typeof api.syncJobsheetOutputs === 'function') {
        try {
          const syncResult = await api.syncJobsheetOutputs({
            businessId: numericBusinessId,
            jobsheetId: normalizedJobsheetId,
            jobsheetSnapshot: currentState,
            hintPaths: filtered.map(doc => doc?.file_path).filter(Boolean)
          });

          if (syncResult?.added > 0) {
            const refreshedDocs = await fetchDocuments();
            filtered = filterForJobsheet(refreshedDocs);

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
    } finally {
      setJobsheetDocumentsLoading(false);
    }
  }, [jobsheetId, numericBusinessId]);

  useEffect(() => {
    if (!jobsheetId) {
      setJobsheetDocuments([]);
      setJobsheetDocumentsError('');
      setJobsheetDocumentsLoading(false);
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

  const handleDeleteJobsheetDocument = useCallback(async (doc) => {
    if (!DOCUMENT_GENERATION_ENABLED && !DOCUMENT_FEATURES_ENABLED) {
      setJobsheetDocumentsError('Document access is currently disabled.');
      return;
    }
    if (!doc || doc.document_id == null) return;
    const typeLabel = DOCUMENT_TYPE_LABELS[doc.doc_type] || startCaseKey(doc.doc_type || 'document');
    const title = typeLabel
      ? `${typeLabel}${doc.number ? ` #${doc.number}` : ''}`
      : 'this document';
    const confirmed = window.confirm(`Delete ${title}? This removes it from this jobsheet.`);
    if (!confirmed) return;

    let removeFile = false;
    if (doc.file_path) {
      removeFile = window.confirm('Also remove the generated file from disk?');
    }

    try {
      setJobsheetDocumentsError('');
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
      requires_total: definition.requires_total ? 1 : 0,
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
        case 'file_suffix':
          next.file_suffix = value;
          break;
        case 'template_path':
          next.template_path = value || '';
          break;
        case 'requires_total':
          next.requires_total = value ? 1 : 0;
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
      file_suffix: definitionDraft.file_suffix ? String(definitionDraft.file_suffix) : null,
      invoice_variant: docType === 'invoice' && definitionDraft.invoice_variant ? String(definitionDraft.invoice_variant) : null,
      template_path: definitionDraft.template_path ? String(definitionDraft.template_path) : null,
      requires_total: definitionDraft.requires_total ? 1 : 0,
      is_primary: definitionDraft.is_primary ? 1 : 0,
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
    creatingRef.current = nextTarget == null;
    setError('');
    setMessage('');
    if (nextTarget != null) {
      const storedSection = getStoredSection(nextTarget);
      const fallbackSection = storedSection || activeEditorSection || 'client';
      setActiveEditorSection(fallbackSection);
    } else {
      setActiveEditorSection('client');
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
  }, [isInline, targetJobsheetId, jobsheetId, numericBusinessId, getStoredSection, activeEditorSection]);

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
        const fallbackSection = storedSection || activeEditorSection || 'client';
        setActiveEditorSection(fallbackSection);
      } else {
        setActiveEditorSection('client');
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
  }, [isInline, numericBusinessId, jobsheetId, getStoredSection, activeEditorSection]);

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
          }
        } else {
          setFormState(DEFAULT_JOBSHEET(numericBusinessId));
          setActiveEditorSection('client');
        }
        initialLoadRef.current = true;
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
  }, [numericBusinessId, initialJobsheetId, jobsheetId, isInline, getStoredSection, activeEditorSection]);

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
    autoSaveTimer.current = setTimeout(async () => {
      setSaving(true);
      try {
        const payload = preparePayload(formState, numericBusinessId);
        await api.updateAhmenJobsheet(jobsheetId, payload);
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
    setSaving(true);
    try {
      const payload = preparePayload(currentState, numericBusinessId);
      await api.updateAhmenJobsheet(jobsheetId, payload);
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

  useEffect(() => {
    if (loading || !jobsheetId) return;
    if (initialLoadRef.current) {
      initialLoadRef.current = false;
      return;
    }
    saveJobsheet();
  }, [formState, jobsheetId, loading, saveJobsheet]);

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

    if (definition.file_suffix) payload.file_name_suffix = definition.file_suffix;
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

    const total = parseAmount(current.pricing_total)
      ?? (pricingDerived ? parseAmount(pricingDerived.total) : null)
      ?? parseAmount(current.ahmen_fee);

    const docType = (definition?.doc_type || '').toLowerCase();
    const requiresTotal = definition?.requires_total ? true : false;
    const needsTotal = requiresTotal || docType === 'invoice' || docType === 'quote' || docType === 'workbook';

    if (needsTotal && !total) {
      messages.push('Enter at least one fee before generating.');
    }

    if (definition && definition.is_active === 0) {
      messages.push('This document type is inactive. Reactivate it before generating.');
    }

    return messages;
  }, [numericBusinessId, parseAmount, pricingDerived]);

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

    const meta = DOC_TYPE_META[definition.doc_type] || null;

    setDocumentGeneratingKey(definition.key);
    setError('');
    try {
      if (templatePath && meta?.supportsNormalize && templatePath.toLowerCase().endsWith('.xlsx')) {
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
      || jobsheetDocuments.find(doc => (doc.definition_key || 'workbook') === targetKey || (doc.doc_type === 'workbook' && !doc.definition_key));

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
    let folderPath = resolvedBusiness?.save_path;
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
  }, [resolvedBusiness, numericBusinessId, setBusiness]);

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
    if (loading) return;
    if (jobsheetId) return;
    if (!formState.client_name?.trim()) return;
    if (creatingRef.current) return;
    creatingRef.current = true;
    (async () => {
      const api = window.api;
      if (!api || !api.addAhmenJobsheet) return;
      try {
        setSaving(true);
        const payload = preparePayload(formState, numericBusinessId);
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
          snapshot: buildSnapshot({ ...formState, jobsheet_id: newId }, newId)
        });
        setMessage('Draft created');
        setTimeout(() => setMessage(''), 1500);
        initialLoadRef.current = true;
      } catch (err) {
        console.error('Failed to create jobsheet', err);
        setError(err?.message || 'Unable to create jobsheet');
      } finally {
        creatingRef.current = false;
        setSaving(false);
      }
    })();
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
      if (!api || !api.deleteAhmenJobsheet) {
        setError('Unable to delete jobsheet: API unavailable');
        setSaving(false);
        return;
      }
      await api.deleteAhmenJobsheet(jobsheetId);
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
        summaryCard
      ) : (
        <div className="sticky top-0 z-20 -mx-6 px-6 pt-2 pb-4 bg-slate-100/95 backdrop-blur">
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
      />
      <JobsheetDocumentsPanel
        jobsheetId={jobsheetId}
        documents={jobsheetDocuments}
        loading={jobsheetDocumentsLoading}
        definitionsLoading={documentDefinitionsLoading}
        error={jobsheetDocumentsError}
        onRefresh={refreshJobsheetDocuments}
        onGenerate={handlePopulateExcel}
        onRegenerate={handleRegenerateWorkbook}
        generatingKey={documentGeneratingKey}
        workbookDefinition={workbookDefinition}
        onOpenTemplate={handleOpenDefinitionTemplate}
        onEditTemplate={openEditDefinitionModal}
        onOpenOutputFolder={handleOpenOutputFolder}
        onOpenOutputFile={handleOpenOutputFile}
        onOpenFile={handleOpenDocumentFile}
        onRevealFile={handleRevealDocument}
        onDelete={handleDeleteJobsheetDocument}
        lastOutputPath={lastOutputPath}
      />
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
              <label className="block text-sm font-medium text-slate-600">
                File suffix (optional)
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={definitionDraft.file_suffix}
                  onChange={event => handleDefinitionDraftChange('file_suffix', event.target.value)}
                  placeholder="e.g. - Statement of Work"
                />
              </label>
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
                  checked={Boolean(definitionDraft.requires_total)}
                  onChange={event => handleDefinitionDraftChange('requires_total', event.target.checked)}
                />
                Requires total amount
              </label>
              <label className="inline-flex items-center gap-2">
                <input
                  type="checkbox"
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                  checked={Boolean(definitionDraft.is_primary)}
                  onChange={event => handleDefinitionDraftChange('is_primary', event.target.checked)}
                />
                Mark as primary option
              </label>
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

  if (isInline) {
    const inlineStatus = saving ? 'Saving…' : message;
    const inlineMessageVisible = !error && Boolean(inlineStatus);
    const inlineDisplay = inlineStatus || '\u00A0';
    return (
      <div className="space-y-4 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 xl:px-10 py-4 sm:py-6">
        {error ? <div className="rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
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
        {error ? <div className="rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
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
        const businessList = data || [];
        setBusinesses(businessList);

        const storedId = readLastBusinessId();
        if (!selectedBusiness && storedId) {
          const match = businessList.find(biz => String(biz.id) === storedId);
          if (match) {
            setSelectedBusiness(match);
          }
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
      <BusinessChooser
        businesses={businesses}
        loading={loading}
        error={error}
        onSelect={handleSelectBusiness}
      />
    );
  }

  return (
    <BusinessWorkspace
      business={selectedBusiness}
      onSwitch={() => setSelectedBusiness(null)}
      onBusinessUpdate={handleBusinessUpdated}
    />
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = createRoot(rootElement);
  root.render(<App />);
}

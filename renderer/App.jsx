import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';

const AHMEN_NUMERIC_FIELDS = new Set([
  'ahmen_fee',
  'specialist_fees',
  'production_fees',
  'deposit_amount',
  'balance_amount',
  'pricing_discount',
  'pricing_total'
]);

const AHMEN_BOOLEAN_FIELDS = new Set(['venue_same_as_client']);

const STATUS_OPTIONS = [
  { value: 'enquiry', label: 'Enquiry' },
  { value: 'quoted', label: 'Quoted' },
  { value: 'confirmed', label: 'Confirmed' },
  { value: 'completed', label: 'Completed' }
];

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

const STATUS_ORDER = STATUS_OPTIONS.reduce((acc, option, index) => {
  acc[option.value] = index;
  return acc;
}, {});

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
  { key: 'client_name', label: 'Client', sortable: true },
  { key: 'event_type', label: 'Event', sortable: true },
  { key: 'event_date', label: 'Event Date', sortable: true },
  { key: 'status', label: 'Status', sortable: true },
  { key: 'ahmen_fee', label: 'Fee', sortable: true, align: 'right' },
  { key: 'actions', label: '', sortable: false, align: 'right' }
];

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
  specialist_fees: '',
  production_fees: '',
  deposit_amount: '',
  balance_amount: '',
  balance_due_date: '',
  balance_reminder_date: '',
  service_types: '',
  specialist_singers: '',
  notes: '',
  pricing_service_id: '',
  pricing_selected_singers: [],
  pricing_custom_fees: '',
  pricing_discount: '',
  pricing_total: ''
});

const FORM_GROUPS = [
  {
    key: 'client',
    title: 'Client Details',
    description: 'Captured during the initial enquiry.',
    defaultOpen: true,
    fields: [
      {
        name: 'status',
        label: 'Status',
        component: 'statusSelect',
        options: STATUS_OPTIONS
      },
      { name: 'client_name', label: 'Client Name', required: true },
      { name: 'client_email', label: 'Email', type: 'email' },
      { name: 'client_phone', label: 'Phone' },
      { name: 'client_address1', label: 'Address Line 1' },
      { name: 'client_address2', label: 'Address Line 2' },
      { name: 'client_address3', label: 'Address Line 3' },
      { name: 'client_town', label: 'Town / City' },
      { name: 'client_postcode', label: 'Postcode' }
    ]
  },
  {
    key: 'event',
    title: 'Event Details',
    description: 'What, when, and how the event will run.',
    defaultOpen: false,
    fields: [
      { name: 'event_type', label: 'Event Type' },
      { name: 'event_date', label: 'Event Date', type: 'date' },
      { name: 'event_start', label: 'Start Time', type: 'time' },
      { name: 'event_end', label: 'End Time', type: 'time' }
    ]
  },
  {
    key: 'venue',
    title: 'Venue Details',
    description: 'Where your team will be performing and saved venue options.',
    defaultOpen: false,
    fields: [
      { name: 'saved_venue', label: 'Saved Venue', component: 'savedVenueSelector' },
      { name: 'venue_same_as_client', label: 'Use client address (private residence)', type: 'checkbox', hint: 'Copies the client address and does not save the venue to the shared directory.' },
      { name: 'venue_name', label: 'Venue Name' },
      { name: 'venue_address1', label: 'Address Line 1' },
      { name: 'venue_address2', label: 'Address Line 2' },
      { name: 'venue_address3', label: 'Address Line 3' },
      { name: 'venue_town', label: 'Town / City' },
      { name: 'venue_postcode', label: 'Postcode' }
    ]
  },
  {
    key: 'billing',
    title: 'Billing Details',
    description: 'Financial breakdown that feeds quotes and invoices.',
    defaultOpen: false,
    fields: [
      { name: 'ahmen_fee', label: 'AhMen Fee (£)', type: 'number', step: '0.01', hint: 'Total fee for the booking.' },
      { name: 'specialist_fees', label: 'Specialist Singers / Other Fees (£)', type: 'number', step: '0.01' },
      { name: 'production_fees', label: 'Sound / AV / Production (£)', type: 'number', step: '0.01' },
      { name: 'deposit_amount', label: 'Deposit (£)', type: 'number', step: '0.01', readOnly: true, hint: 'Automatically 30% of AhMen fee.' },
      { name: 'balance_amount', label: 'Balance (£)', type: 'number', step: '0.01', readOnly: true, hint: 'Remaining balance after deposit (70%).' },
      { name: 'balance_due_date', label: 'Balance Due Date', type: 'date', readOnly: true, hint: 'Automatically 10 days before the event.' },
      { name: 'balance_reminder_date', label: 'Balance Reminder Date', type: 'date', readOnly: true, hint: 'Automatically 20 days before the event.' }
    ]
  },
  {
    key: 'services',
    title: 'Services & Notes',
    description: 'Additional requirements and context for the booking.',
    defaultOpen: false,
    fields: [
      { name: 'service_types', label: 'Service Type(s)', type: 'textarea', rows: 2 },
      { name: 'specialist_singers', label: 'Specialist Singers', type: 'textarea', rows: 2 },
      { name: 'notes', label: 'Internal Notes', type: 'textarea', rows: 3 }
    ]
  }
];

function formatDateInput(value) {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return value.toISOString().slice(0, 10);
  }
  const date = new Date(value);
  if (Number.isNaN(date.valueOf())) return '';
  return date.toISOString().slice(0, 10);
}

function formatDateDisplay(value) {
  if (!value) return 'Date tbc';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.valueOf())) return 'Date tbc';
  return parsed.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'long',
    year: 'numeric'
  });
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
  if (Array.isArray(raw)) return raw;
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function preparePayload(formState, businessId) {
  const payload = { ...formState, business_id: businessId };

  if (Array.isArray(payload.pricing_selected_singers)) {
    payload.pricing_selected_singers = JSON.stringify(payload.pricing_selected_singers);
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

  const fee = Number(next.ahmen_fee);
  if (Number.isFinite(fee) && fee >= 0) {
    const deposit = Math.round(fee * 0.3 * 100) / 100;
    const balance = Math.max(fee - deposit, 0);
    next.deposit_amount = deposit.toFixed(2);
    next.balance_amount = balance.toFixed(2);
  } else {
    next.deposit_amount = '';
    next.balance_amount = '';
  }

  if (next.event_date) {
    next.balance_due_date = addDays(next.event_date, -10);
    next.balance_reminder_date = addDays(next.event_date, -20);
  } else {
    next.balance_due_date = '';
    next.balance_reminder_date = '';
  }

  if (next.venue_same_as_client) {
    next.venue_name = next.client_name || next.venue_name;
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

function JobsheetList({ jobsheets, selectedId, onSelect, onNew, onDelete, onStatusChange, loading, deleting, sortConfig, onSort }) {
  const sortedJobsheets = useMemo(() => {
    const list = [...jobsheets];
    const { key, direction } = sortConfig || {};
    if (!key) return list;
    const multiplier = direction === 'asc' ? 1 : -1;

    const getComparableValue = (sheet, field) => {
      switch (field) {
        case 'event_date':
          return sheet.event_date ? new Date(sheet.event_date).valueOf() : 0;
        case 'ahmen_fee':
          return Number(sheet.ahmen_fee) || 0;
        case 'status':
          return STATUS_ORDER[sheet.status] ?? STATUS_OPTIONS.length;
        case 'client_name':
        case 'event_type':
          return (sheet[field] || '').toString().toLowerCase();
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
  }, [jobsheets, sortConfig]);

  const renderSortIndicator = (columnKey) => {
    if (!sortConfig || sortConfig.key !== columnKey) return <span className="text-slate-400 ml-1">⇅</span>;
    return (
      <span className="ml-1 text-xs text-indigo-600">
        {sortConfig.direction === 'asc' ? '▲' : '▼'}
      </span>
    );
  };

  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-lg font-semibold text-slate-700">Jobsheets</h2>
        <button
          onClick={onNew}
          className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-medium px-3 py-2 rounded"
        >
          + New Jobsheet
        </button>
      </div>
      <div className="flex-1 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {loading ? (
          <div className="p-6 text-center text-slate-500">Loading…</div>
        ) : !sortedJobsheets.length ? (
          <div className="p-6 text-center text-slate-500">No jobsheets yet. Create your first one!</div>
        ) : (
          <div className="overflow-y-auto">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-50">
                <tr>
                  {JOBSHEET_COLUMNS.map(column => {
                    const alignment = column.align === 'right' ? 'text-right' : column.align === 'center' ? 'text-center' : 'text-left';
                    if (!column.sortable) {
                      return (
                        <th
                          key={column.key}
                          scope="col"
                          className={`px-4 py-3 font-semibold uppercase tracking-wide text-xs text-slate-500 ${alignment}`}
                        >
                          {column.label}
                        </th>
                      );
                    }
                    return (
                      <th key={column.key} scope="col" className={`px-4 py-3 font-semibold uppercase tracking-wide text-xs text-slate-500 ${alignment}`}>
                        <button
                          type="button"
                          onClick={() => onSort?.(column.key)}
                          className="inline-flex items-center gap-1 text-slate-600 hover:text-indigo-600"
                        >
                          {column.label}
                          {renderSortIndicator(column.key)}
                        </button>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {sortedJobsheets.map(sheet => {
                  const isActive = sheet.jobsheet_id === selectedId;
                  const statusKey = normalizeStatus(sheet.status) || 'enquiry';
                  const statusOption = STATUS_OPTIONS.find(opt => opt.value === statusKey);
                  const statusStyles = STATUS_STYLES[statusKey] || 'bg-slate-200 text-slate-700 border border-slate-300';
                  const statusRowClass = STATUS_ROW_CLASSES[statusKey] || 'bg-white';
                  return (
                    <tr
                      key={sheet.jobsheet_id || sheet.client_name}
                      onClick={() => onSelect(sheet.jobsheet_id)}
                      className={`${statusRowClass} ${isActive ? 'ring-2 ring-indigo-400 ring-inset' : 'hover:shadow-sm'} cursor-pointer transition`}
                    >
                      <td className="px-4 py-3 text-sm font-medium text-slate-800 whitespace-nowrap">
                        {sheet.client_name || 'Untitled booking'}
                      </td>
                      <td className="px-4 py-3 text-sm text-slate-600">{sheet.event_type || '—'}</td>
                      <td className="px-4 py-3 text-sm text-slate-600 whitespace-nowrap">{formatDateDisplay(sheet.event_date)}</td>
                      <td className="px-4 py-3">
                        <select
                          value={statusKey}
                          onClick={event => event.stopPropagation()}
                          onChange={event => onStatusChange?.(sheet.jobsheet_id, event.target.value)}
                          className={`rounded-full border border-transparent px-3 py-1 text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-400 ${statusStyles}`}
                        >
                          {STATUS_OPTIONS.map(option => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td className="px-4 py-3 text-right text-sm text-slate-600">{toCurrency(sheet.ahmen_fee)}</td>
                      <td className="px-4 py-3 text-right text-sm">
                        <div className="inline-flex items-center gap-2">
                          <button
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation();
                              onSelect(sheet.jobsheet_id);
                            }}
                            className="rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100"
                          >
                            Edit
                          </button>
                          <button
                            type="button"
                            disabled={deleting}
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
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function PricingPanel({ pricingConfig, formState, onChange, pricingTotals }) {
  const serviceTypes = pricingConfig?.serviceTypes ?? [];
  const selectedService = serviceTypes.find(type => type.id === formState.pricing_service_id);
  const selectedSingerSet = new Set(formState.pricing_selected_singers || []);

  const internalTotals = useMemo(() => {
    if (!selectedService) return { base: 0, singerCount: 0 };
    let base = 0;
    let singerCount = 0;
    selectedService.singers.forEach(singer => {
      if (selectedSingerSet.has(singer.id)) {
        base += Number(singer.fee) || 0;
        singerCount += 1;
      }
    });
    return { base, singerCount };
  }, [selectedService, selectedSingerSet]);

  const totals = pricingTotals || internalTotals;

  return (
    <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-4">
      <div>
        <h3 className="text-base font-semibold text-slate-700">Pricing</h3>
        <p className="text-sm text-slate-500">Build a quote straight from the pricing template.</p>
      </div>

      <label className="block text-sm font-medium text-slate-600">
        Service configuration
        <select
          className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm"
          value={formState.pricing_service_id || ''}
          onChange={event => onChange('pricing_service_id', event.target.value)}
        >
          <option value="">Select service type…</option>
          {serviceTypes.map(type => (
            <option key={type.id} value={type.id}>{type.label}</option>
          ))}
        </select>
      </label>

      {selectedService ? (
        <div className="space-y-3">
          <div className="rounded border border-slate-200 p-3 max-h-56 overflow-y-auto">
            <div className="text-sm font-medium text-slate-600 mb-2">Singers</div>
            <div className="space-y-2">
              {selectedService.singers.map(singer => {
                const checked = selectedSingerSet.has(singer.id);
                return (
                  <label key={singer.id} className="flex items-start gap-2 text-sm text-slate-600">
                    <input
                      type="checkbox"
                      className="mt-1"
                      checked={checked}
                      onChange={() => {
                        const next = new Set(selectedSingerSet);
                        if (checked) {
                          next.delete(singer.id);
                        } else {
                          next.add(singer.id);
                        }
                        onChange('pricing_selected_singers', Array.from(next));
                      }}
                    />
                    <span>
                      <span className="font-medium text-slate-700">{singer.name}</span>
                      <span className="block text-xs text-slate-500">Fee: {toCurrency(singer.fee)}{singer.comments ? ` · ${singer.comments}` : ''}</span>
                    </span>
                  </label>
                );
              })}
              {!selectedService.singers.length ? (
                <div className="text-sm text-slate-500">No singers configured in the pricing template.</div>
              ) : null}
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <Field
              label="Custom fees (£)"
              type="number"
              step="0.01"
              value={formState.pricing_custom_fees || ''}
              onChange={value => onChange('pricing_custom_fees', value)}
            />
            <Field
              label="Discount (£)"
              type="number"
              step="0.01"
              value={formState.pricing_discount || ''}
              onChange={value => onChange('pricing_discount', value)}
            />
            <Field
              label="Pricing total (£)"
              type="number"
              step="0.01"
              value={formState.pricing_total || ''}
              onChange={value => onChange('pricing_total', value)}
              readOnly
            />
          </div>

          <div className="rounded-lg bg-indigo-50 p-3 text-sm text-indigo-700">
            <div className="font-semibold">Quote summary</div>
            <div>{totals.singerCount} singers selected · Base fee {toCurrency(totals.base)}</div>
            <div>Total after adjustments: {toCurrency(formState.pricing_total)}</div>
          </div>
        </div>
      ) : (
        <div className="text-sm text-slate-500">Select a service type to see preset singers and fees.</div>
      )}
    </div>
  );
}

function Field({ label, type = 'text', value, onChange, readOnly, hint, rows = 3, step, component, options, secondaryAction, secondaryDisabled }) {
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
    input = (
      <select
        className='mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500'
        value={value || 'enquiry'}
        onChange={event => onChange(event.target.value)}
      >
        {(options || STATUS_OPTIONS).map(option => (
          <option key={option.value} value={option.value}>{option.label}</option>
        ))}
      </select>
    );
  } else if (component === 'savedVenueSelector') {
    input = (
      <div className="space-y-2">
        <select
          className='mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500'
          value={value || ''}
          onChange={event => onChange(event.target.value)}
        >
          <option value="">Select saved venue…</option>
          {options?.map(venue => (
            <option key={venue.venue_id} value={venue.venue_id}>
              {venue.name || 'Untitled venue'}
            </option>
          ))}
        </select>
        <button
          type="button"
          onClick={() => secondaryAction?.('SAVE_CURRENT_VENUE')}
          disabled={secondaryDisabled}
          className="inline-flex items-center rounded bg-slate-800 px-3 py-2 text-xs font-medium text-white hover:bg-slate-700 disabled:opacity-60 disabled:cursor-not-allowed"
        >
          {secondaryDisabled ? 'Saving…' : 'Save current venue'}
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

function JobsheetEditor({
  business,
  formState,
  onChange,
  onSave,
  onDelete,
  saving,
  deleting,
  hasExisting,
  venues,
  onSaveVenue,
  venueSaving,
  pricingConfig,
  pricingTotals
}) {
  const handleFieldChange = (name, value) => {
    const next = applyDerivedFields({ ...formState, [name]: value });
    onChange(next);
  };

  const [savedVenueId, setSavedVenueId] = useState(() => (
    formState.venue_id ? String(formState.venue_id) : ''
  ));

  useEffect(() => {
    setSavedVenueId(formState.venue_id ? String(formState.venue_id) : '');
  }, [formState.venue_id]);

  const [openGroups, setOpenGroups] = useState(() => {
    const initial = {};
    FORM_GROUPS.forEach(group => {
      initial[group.key] = group.defaultOpen ?? false;
    });
    return initial;
  });

  const toggleGroup = (key) => {
    setOpenGroups(prev => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  return (
    <div className="space-y-6">
      <div className="flex items-start justify-between">
        <div>
          <h2 className="text-xl font-semibold text-slate-800">{hasExisting ? 'Edit jobsheet' : 'New jobsheet'}</h2>
          <p className="text-sm text-slate-500">Business: {business.business_name}</p>
        </div>
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

      <div className="space-y-4">
        {FORM_GROUPS.map(group => (
          <section key={group.key} className="bg-white border border-slate-200 rounded-lg">
            <button
              type="button"
              onClick={() => toggleGroup(group.key)}
              className="w-full flex items-center justify-between px-5 py-4"
            >
              <div className="text-left">
                <h3 className="text-base font-semibold text-slate-700">{group.title}</h3>
                <p className="text-sm text-slate-500">{group.description}</p>
              </div>
              <span className="text-xl text-slate-500">{openGroups[group.key] ? '−' : '+'}</span>
            </button>
            {openGroups[group.key] ? (
              <div className="border-t border-slate-200 px-5 py-4 space-y-4">
                {group.fields.map(field => {
                  const resolvedValue = field.component === 'savedVenueSelector'
                    ? savedVenueId
                    : field.name === 'status'
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
                      readOnly={field.readOnly}
                      component={field.component}
                      options={field.component === 'savedVenueSelector' ? venues : field.options}
                      value={resolvedValue}
                      onChange={value => {
                        if (field.component === 'savedVenueSelector') {
                          setSavedVenueId(value || '');
                          if (!value) {
                            handleFieldChange('venue_id', null);
                            return;
                          }
                          const venue = venues.find(v => String(v.venue_id) === value);
                          if (venue) {
                            handleFieldChange('venue_id', venue.venue_id);
                            handleFieldChange('venue_name', venue.name || '');
                            handleFieldChange('venue_address1', venue.address1 || '');
                            handleFieldChange('venue_address2', venue.address2 || '');
                            handleFieldChange('venue_address3', venue.address3 || '');
                            handleFieldChange('venue_town', venue.town || '');
                            handleFieldChange('venue_postcode', venue.postcode || '');
                          }
                          return;
                        }
                        handleFieldChange(
                          field.name,
                          field.type === 'checkbox' ? Boolean(value) : value
                        );
                      }}
                      secondaryAction={action => {
                        if (field.component === 'savedVenueSelector' && action === 'SAVE_CURRENT_VENUE') {
                          onSaveVenue();
                        }
                      }}
                      secondaryDisabled={field.component === 'savedVenueSelector' && venueSaving}
                    />
                  );
                })}
              </div>
            ) : null}
          </section>
        ))}
      </div>

      <div className="grid gap-6 lg:grid-cols-2">
        <PricingPanel
          pricingConfig={pricingConfig}
          pricingTotals={pricingTotals}
          formState={formState}
          onChange={(field, value) => handleFieldChange(field, value)}
        />
      </div>

      <div className="flex items-center justify-end gap-3">
        <button
          onClick={onSave}
          disabled={saving}
          className="inline-flex items-center justify-center rounded bg-indigo-600 text-white text-sm font-medium px-4 py-2 hover:bg-indigo-500 disabled:opacity-60"
        >
          {saving ? 'Saving…' : 'Save jobsheet'}
        </button>
      </div>
    </div>
  );
}

function BusinessWorkspace({ business, onSwitch }) {
  const [jobsheets, setJobsheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [listLoading, setListLoading] = useState(true);
  const [formState, setFormState] = useState(DEFAULT_JOBSHEET(business.id));
  const [selectedId, setSelectedId] = useState(null);
  const [isCreating, setIsCreating] = useState(false);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [sortConfig, setSortConfig] = useState({ key: 'event_date', direction: 'desc' });
  const [venues, setVenues] = useState([]);
  const [venueSaving, setVenueSaving] = useState(false);
  const [pricingConfig, setPricingConfig] = useState(null);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    let mounted = true;
    setLoading(true);
    setListLoading(true);
    setError('');
    const loadAll = async () => {
      try {
        const [jobsheetData, venueData, pricingData] = await Promise.all([
          window.api.getAhmenJobsheets({ businessId: business.id }),
          window.api.getAhmenVenues({ businessId: business.id }),
          window.api.getAhmenPricing()
        ]);
        if (!mounted) return;
        setJobsheets((jobsheetData || []).map(item => ({
          ...item,
          status: normalizeStatus(item.status) || 'enquiry'
        })));
        setVenues((venueData || []).map(item => ({
          ...item,
          venue_id: item.venue_id ?? item.id,
          name: item.name || item.venue_name || '',
          address1: item.address1 || item.venue_address1 || '',
          address2: item.address2 || item.venue_address2 || '',
          address3: item.address3 || item.venue_address3 || '',
          town: item.town || item.venue_town || '',
          postcode: item.postcode || item.venue_postcode || ''
        })));
        setPricingConfig(pricingData || null);
      } catch (err) {
        if (!mounted) return;
        console.error('Failed to load business workspace', err);
        setError(err?.message || 'Unable to load business data');
      } finally {
        if (mounted) {
          setLoading(false);
          setListLoading(false);
          setFormState(DEFAULT_JOBSHEET(business.id));
          setSelectedId(null);
          setIsCreating(false);
        }
      }
    };
    loadAll();
    return () => {
      mounted = false;
    };
  }, [business.id]);

  const refreshJobsheets = async () => {
    setListLoading(true);
    try {
      const data = await window.api.getAhmenJobsheets({ businessId: business.id });
      setJobsheets((data || []).map(item => ({
        ...item,
        status: normalizeStatus(item.status) || 'enquiry'
      })));
    } catch (err) {
      console.error('Failed to refresh jobsheets', err);
      setError(err?.message || 'Unable to refresh jobsheets');
    } finally {
      setListLoading(false);
    }
  };

  const handleSelect = async (jobsheetId) => {
    if (!jobsheetId) return;
    try {
      setSaving(false);
      setDeleting(false);
      setIsCreating(false);
      setSelectedId(jobsheetId);
      const sheet = jobsheets.find(item => item.jobsheet_id === jobsheetId);
      if (!sheet) {
        const fetched = await window.api.getAhmenJobsheet(jobsheetId);
        if (fetched) {
          setFormState(mapApiToForm(fetched, business.id));
        }
        return;
      }
      setFormState(mapApiToForm(sheet, business.id));
    } catch (err) {
      console.error('Failed to load jobsheet', err);
      setError(err?.message || 'Unable to load jobsheet');
    }
  };

  const handleNew = () => {
    setSelectedId(null);
    setIsCreating(true);
    setFormState(DEFAULT_JOBSHEET(business.id));
    setError('');
    setMessage('');
  };

  const handleSave = async () => {
    if (!formState.client_name?.trim()) {
      setError('Client name is required');
      return;
    }
    setSaving(true);
    setError('');
    setMessage('');
    try {
      if (selectedId) {
        await window.api.updateAhmenJobsheet(selectedId, preparePayload(formState, business.id));
        setMessage('Jobsheet updated');
      } else {
        const newId = await window.api.addAhmenJobsheet(preparePayload(formState, business.id));
        setMessage('Jobsheet created');
        setSelectedId(newId);
        setFormState(prev => ({ ...prev, jobsheet_id: newId }));
        setIsCreating(false);
      }
      await refreshJobsheets();
      if (selectedId) {
        setIsCreating(false);
      }
    } catch (err) {
      console.error('Failed to save jobsheet', err);
      setError(err?.message || 'Unable to save jobsheet');
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (jobsheetId) => {
    const targetId = jobsheetId ?? selectedId;
    if (!targetId) return;
    const confirmed = window.confirm('Delete this jobsheet? This cannot be undone.');
    if (!confirmed) return;
    setDeleting(true);
    setError('');
    setMessage('');
    try {
      await window.api.deleteAhmenJobsheet(targetId);
      setMessage('Jobsheet deleted');
      await refreshJobsheets();
      if (targetId === selectedId) {
        setSelectedId(null);
        setIsCreating(false);
        setFormState(DEFAULT_JOBSHEET(business.id));
      }
    } catch (err) {
      console.error('Failed to delete jobsheet', err);
      setError(err?.message || 'Unable to delete jobsheet');
    } finally {
      setDeleting(false);
    }
  };

  const handleStatusChange = async (jobsheetId, nextStatus) => {
    if (!jobsheetId || !nextStatus) return;
    setError('');
    setMessage('');
    try {
      let sheet = jobsheets.find(item => item.jobsheet_id === jobsheetId);
      if (!sheet) {
        sheet = await window.api.getAhmenJobsheet(jobsheetId);
      }
      if (!sheet) return;

      const updatedSheet = { ...sheet, status: nextStatus };
      const form = mapApiToForm(updatedSheet, business.id);
      form.status = nextStatus;
      await window.api.updateAhmenJobsheet(jobsheetId, preparePayload(form, business.id));

      setJobsheets(prev => prev.map(item => (
        item.jobsheet_id === jobsheetId ? { ...item, status: nextStatus, updated_at: new Date().toISOString() } : item
      )));

      if (selectedId === jobsheetId) {
        setFormState(prev => ({ ...prev, status: nextStatus }));
      }
      setMessage('Status updated');
    } catch (err) {
      console.error('Failed to update status', err);
      setError(err?.message || 'Unable to update status');
    }
  };

  const handleSort = (columnKey) => {
    if (!columnKey) return;
    setSortConfig(prev => {
      if (prev.key === columnKey) {
        const nextDirection = prev.direction === 'asc' ? 'desc' : 'asc';
        return { key: columnKey, direction: nextDirection };
      }
      return { key: columnKey, direction: columnKey === 'client_name' ? 'asc' : 'desc' };
    });
  };

  const handleSaveVenue = async () => {
    setVenueSaving(true);
    setError('');
    try {
      const venuePayload = {
        business_id: business.id,
        name: formState.venue_name,
        address1: formState.venue_address1,
        address2: formState.venue_address2,
        address3: formState.venue_address3,
        town: formState.venue_town,
        postcode: formState.venue_postcode,
        is_private: formState.venue_same_as_client ? 1 : 0
      };
      if (!venuePayload.name?.trim()) {
        setError('Venue name is required to save.');
        return;
      }
      await window.api.saveAhmenVenue(venuePayload);
      const updatedVenues = await window.api.getAhmenVenues({ businessId: business.id });
      setVenues((updatedVenues || []).map(item => ({
        ...item,
        venue_id: item.venue_id ?? item.id,
        name: item.name || item.venue_name || '',
        address1: item.address1 || item.venue_address1 || '',
        address2: item.address2 || item.venue_address2 || '',
        address3: item.address3 || item.venue_address3 || '',
        town: item.town || item.venue_town || '',
        postcode: item.postcode || item.venue_postcode || ''
      })));
      setMessage('Venue saved');
    } catch (err) {
      console.error('Failed to save venue', err);
      setError(err?.message || 'Unable to save venue');
    } finally {
      setVenueSaving(false);
    }
  };

  const pricingDerived = useMemo(() => {
    if (!pricingConfig) return null;
    const service = pricingConfig.serviceTypes?.find(type => type.id === formState.pricing_service_id);
    const selected = new Set(formState.pricing_selected_singers || []);
    let base = 0;
    let singerCount = 0;
    if (service) {
      service.singers.forEach(singer => {
        if (selected.has(singer.id)) {
          base += Number(singer.fee) || 0;
          singerCount += 1;
        }
      });
    }
    const custom = Number(formState.pricing_custom_fees) || 0;
    const discount = Number(formState.pricing_discount) || 0;
    const hasSelection = singerCount > 0 || custom !== 0 || discount !== 0;
    const total = Math.max(base + custom - discount, 0);
    const totalString = hasSelection ? total.toFixed(2) : '';
    return {
      base,
      singerCount,
      hasSelection,
      total,
      totalString
    };
  }, [pricingConfig, formState.pricing_service_id, formState.pricing_selected_singers, formState.pricing_custom_fees, formState.pricing_discount]);

  useEffect(() => {
    if (!pricingDerived) return;
    const { totalString, hasSelection } = pricingDerived;
    if (!hasSelection && !formState.pricing_total && !formState.ahmen_fee) return;
    setFormState(prev => {
      const currentTotal = prev.pricing_total ?? '';
      const currentFee = prev.ahmen_fee ?? '';
      const nextTotal = totalString;
      const shouldUpdateTotal = nextTotal !== currentTotal;
      const matchesPrevTotal = currentFee === currentTotal || !currentFee;
      const shouldUpdateFee = shouldUpdateTotal && matchesPrevTotal;
      if (!shouldUpdateTotal && !shouldUpdateFee) return prev;
      const next = { ...prev };
      next.pricing_total = nextTotal;
      if (shouldUpdateFee) {
        next.ahmen_fee = nextTotal;
      }
      return applyDerivedFields(next);
    });
  }, [pricingDerived]);

  return (
    <div className="min-h-screen bg-slate-100">
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold text-slate-800">{business.business_name}</h1>
            <p className="text-sm text-slate-500">Manage AhMen jobsheets, venues, and pricing.</p>
          </div>
          <button
            onClick={onSwitch}
            className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
          >
            Switch business
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6">
        {loading ? (
          <div className="bg-white rounded-lg border border-slate-200 p-6 text-center text-slate-500">Loading workspace…</div>
        ) : (
          <div className="grid gap-6 lg:grid-cols-[320px_1fr]">
            <JobsheetList
              jobsheets={jobsheets}
              selectedId={selectedId}
              onSelect={handleSelect}
              onNew={handleNew}
              onDelete={handleDelete}
              onStatusChange={handleStatusChange}
              loading={listLoading}
              deleting={deleting}
              sortConfig={sortConfig}
              onSort={handleSort}
            />
            <div>
              {error ? <div className="mb-4 rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
              {message ? <div className="mb-4 rounded border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-700">{message}</div> : null}
              {(isCreating || selectedId !== null) ? (
                <JobsheetEditor
                  business={business}
                  formState={formState}
                  onChange={setFormState}
                  onSave={handleSave}
                  onDelete={() => handleDelete()}
                  saving={saving}
                  deleting={deleting}
                  hasExisting={Boolean(selectedId)}
                  venues={venues}
                  onSaveVenue={handleSaveVenue}
                  venueSaving={venueSaving}
                  pricingConfig={pricingConfig}
                  pricingTotals={pricingDerived}
                />
              ) : (
                <div className="rounded-lg border border-dashed border-slate-300 bg-white px-6 py-16 text-center text-slate-500">
                  <p className="text-lg font-semibold text-slate-600">Select a jobsheet to view</p>
                  <p className="mt-2 text-sm">Choose one from the list on the left or create a new jobsheet to begin.</p>
                  <div className="mt-4">
                    <button
                      onClick={handleNew}
                      className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500"
                    >
                      + New Jobsheet
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
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
    if (key === 'venue_same_as_client') {
      base[key] = Boolean(value);
      return;
    }
    if (key === 'status') {
      const normalized = normalizeStatus(value) || 'enquiry';
      base[key] = normalized;
      return;
    }
    base[key] = value ?? base[key] ?? '';
  });
  return applyDerivedFields(base);
}

function App() {
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
    />
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = createRoot(rootElement);
  root.render(<App />);
}

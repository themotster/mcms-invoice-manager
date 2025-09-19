import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
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

function normalizeVenues(list = []) {
  return (list || []).map(item => ({
    ...item,
    venue_id: item.venue_id ?? item.id,
    name: item.name || item.venue_name || '',
    address1: item.address1 || item.venue_address1 || '',
    address2: item.address2 || item.venue_address2 || '',
    address3: item.address3 || item.venue_address3 || '',
    town: item.town || item.venue_town || '',
    postcode: item.postcode || item.venue_postcode || '',
    is_private: Boolean(item.is_private)
  }));
}

const JOBSHEET_COLUMNS = [
  { key: 'client_name', label: 'Client', sortable: true },
  { key: 'event_type', label: 'Event', sortable: true },
  { key: 'event_date', label: 'Event Date', sortable: true },
  { key: 'venue_name', label: 'Venue', sortable: true },
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
    defaultOpen: false,
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
    key: 'pricing',
    title: 'Pricing & Personnel',
    description: 'Select singers and configure fees for the booking.',
    defaultOpen: false,
    fields: [
      { name: 'pricing_panel', component: 'pricingPanel' }
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
  if (Array.isArray(raw)) return normalizeSingerEntries(raw);
  try {
    const parsed = JSON.parse(raw);
    return normalizeSingerEntries(parsed);
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

  if (next.venue_same_as_client && !next.venue_id) {
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
  onSort
}) {
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
                          onClick={() => onSort(column.key)}
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
                  const statusKey = normalizeStatus(sheet.status) || 'enquiry';
                  const statusStyles = STATUS_STYLES[statusKey] || 'bg-slate-200 text-slate-700 border border-slate-300';
                  const statusDisabled = statusUpdatingId === sheet.jobsheet_id;
                  const statusRowClass = STATUS_ROW_CLASSES[statusKey] || 'bg-white';
                  return (
                    <tr
                      key={sheet.jobsheet_id || sheet.client_name}
                      onClick={() => onOpen(sheet.jobsheet_id)}
                      className={`${statusRowClass} hover:shadow-sm cursor-pointer transition`}
                    >
                      <td className="px-4 py-3 text-sm font-medium text-slate-800 whitespace-nowrap">
                        {sheet.client_name || 'Untitled booking'}
                      </td>
                      <td className="px-4 py-3 text-sm text-slate-600">{sheet.event_type || '—'}</td>
                      <td className="px-4 py-3 text-sm text-slate-600 whitespace-nowrap">{formatDateDisplay(sheet.event_date)}</td>
                      <td className="px-4 py-3 text-sm text-slate-600">
                        {sheet.venue_name || sheet.venue_town || sheet.venue_address1 || '—'}
                      </td>
                      <td className="px-4 py-3">
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
                      </td>
                      <td className="px-4 py-3 text-right text-sm text-slate-600">{toCurrency(sheet.ahmen_fee)}</td>
                      <td className="px-4 py-3 text-right text-sm">
                        <div className="inline-flex items-center">
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
                })}
              </tbody>
            </table>
          </div>
        )}
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
    if (typeof entry === 'string') {
      id = entry;
    } else if (typeof entry === 'object') {
      id = entry.id ?? entry.singerId ?? entry.value;
      if (entry.fee !== undefined && entry.fee !== null) {
        fee = entry.fee === '' ? '' : String(entry.fee);
      }
    }
    if (!id) return;
    const key = String(id);
    if (seen.has(key)) return;
    seen.add(key);
    normalized.push({ id: key, fee });
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
  }
  return true;
}

function PricingPanel({ pricingConfig, formState, onChange, pricingTotals, hasExisting = false }) {
  const serviceTypes = pricingConfig?.serviceTypes ?? [];
  const selectedService = serviceTypes.find(type => type.id === formState.pricing_service_id);

  const selectedEntries = useMemo(
    () => normalizeSingerEntries(formState.pricing_selected_singers),
    [formState.pricing_selected_singers]
  );

  const lastServiceIdRef = useRef(null);
  const initializedRef = useRef(false);
  const previousServiceIdRef = useRef(null);
  const serviceSingers = Array.isArray(selectedService?.singers) ? selectedService.singers : [];

  const updateSelected = useCallback((entries) => {
    const normalized = normalizeSingerEntries(entries);
    if (!equalSingerEntries(normalized, selectedEntries)) {
      onChange('pricing_selected_singers', normalized);
    }
  }, [onChange, selectedEntries]);

  useEffect(() => {
    if (!selectedService) {
      if (selectedEntries.length) {
        updateSelected([]);
      }
      lastServiceIdRef.current = null;
      initializedRef.current = false;
      previousServiceIdRef.current = null;
      return;
    }

    const currentServiceId = selectedService.id;
    const previousServiceId = lastServiceIdRef.current;
    const isServiceChange = previousServiceId !== null && currentServiceId !== previousServiceId;

    if (currentServiceId !== previousServiceId) {
      lastServiceIdRef.current = currentServiceId;
      initializedRef.current = false;
    }

    const serviceMap = new Map(serviceSingers.map(singer => [String(singer.id), singer]));

    let next = selectedEntries
      .filter(entry => serviceMap.has(entry.id))
      .map(entry => {
        const singer = serviceMap.get(entry.id);
        const fee = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
          ? String(entry.fee)
          : singer?.fee != null ? String(singer.fee) : '';
        return { id: entry.id, fee };
      });

    const isFirstVisit = previousServiceIdRef.current === null;
    if (isFirstVisit && hasExisting) {
      previousServiceIdRef.current = currentServiceId;
    } else if (previousServiceIdRef.current !== currentServiceId) {
      serviceSingers.forEach(singer => {
        if (!singer.defaultIncluded) return;
        const singerId = String(singer.id);
        if (!next.find(entry => entry.id === singerId)) {
          next.push({ id: singerId, fee: singer.fee != null ? String(singer.fee) : '' });
        }
      });
      previousServiceIdRef.current = currentServiceId;
    }

    if (!initializedRef.current) {
      initializedRef.current = true;
      if (!selectedEntries.length && !next.length) {
        const shouldSeedDefaults = (!hasExisting && currentServiceId != null) || isServiceChange;
        if (!shouldSeedDefaults) {
          const normalized = normalizeSingerEntries(next);
          if (!equalSingerEntries(normalized, selectedEntries)) {
            updateSelected(normalized);
          }
          return;
        }
        next = serviceSingers
          .filter(singer => singer.defaultIncluded)
          .map(singer => ({ id: String(singer.id), fee: singer.fee != null ? String(singer.fee) : '' }));
      }
    }

    const normalized = normalizeSingerEntries(next);
    if (!equalSingerEntries(normalized, selectedEntries)) {
      updateSelected(normalized);
    }
  }, [selectedService, selectedEntries, updateSelected, serviceSingers, hasExisting]);

  const internalTotals = useMemo(() => {
    if (!selectedService) return { base: 0, singerCount: 0 };
    let base = 0;
    let singerCount = 0;
    serviceSingers.forEach(singer => {
      const singerId = String(singer.id);
      const entry = selectedEntries.find(item => item.id === singerId);
      if (!entry) return;
      const feeValue = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
        ? Number(entry.fee)
        : Number(singer.fee);
      base += Number.isFinite(feeValue) ? feeValue : 0;
      singerCount += 1;
    });
    return { base, singerCount };
  }, [selectedService, selectedEntries]);

  const totals = pricingTotals || internalTotals;

  const handleToggleSinger = (singer, checked) => {
    if (!selectedService) return;
    const singerId = String(singer.id);
    if (checked) {
      if (!selectedEntries.find(entry => entry.id === singerId)) {
        updateSelected([
          ...selectedEntries,
          { id: singerId, fee: singer.fee != null ? String(singer.fee) : '' }
        ]);
      }
    } else {
      updateSelected(selectedEntries.filter(entry => entry.id !== singerId));
    }
  };

  const handleSingerFeeChange = (singer, value) => {
    const singerId = String(singer.id);
    const next = selectedEntries.map(entry => (
      entry.id === singerId ? { ...entry, fee: value } : entry
    ));
    updateSelected(next);
  };

  const renderSingers = () => {
    if (!selectedService) {
      return <div className="text-sm text-slate-500">Select a service type to see preset singers and fees.</div>;
    }

    if (!serviceSingers.length) {
      return <div className="text-sm text-slate-500">No singers configured in the pricing template.</div>;
    }

    return (
      <div className="space-y-3">
        <div className="grid gap-2 sm:grid-cols-2 xl:grid-cols-3">
          {serviceSingers.map(singer => {
            const singerId = String(singer.id);
            const entry = selectedEntries.find(item => item.id === singerId);
            const checked = Boolean(entry);
            const feeInputValue = entry && entry.fee !== undefined && entry.fee !== null
              ? String(entry.fee)
              : singer.fee != null ? String(singer.fee) : '';
            const baseFee = toCurrency(singer.fee);
            const toggleSinger = () => {
              handleToggleSinger(singer, !checked);
            };
            return (
              <div
                key={singer.id}
                role="button"
                tabIndex={0}
                onClick={toggleSinger}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    toggleSinger();
                  }
                }}
                className={`rounded-lg border px-3 py-2 transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${checked ? 'border-indigo-200 bg-indigo-50 shadow-sm' : 'border-slate-200 bg-white hover:border-indigo-200'}`}
              >
                <div className="flex items-start justify-between gap-3">
                  <div className="text-xs">
                    <div className="text-sm font-medium text-slate-700">{singer.name}</div>
                    <div className="text-[11px] text-slate-500">{baseFee}{singer.comments ? ` · ${singer.comments}` : ''}</div>
                  </div>
                  <input
                    type="checkbox"
                    className="mt-0.5 h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                    checked={checked}
                    onClick={event => event.stopPropagation()}
                    onChange={event => handleToggleSinger(singer, event.target.checked)}
                  />
                </div>
                {checked ? (
                  <div className="mt-2 flex items-center gap-2">
                    <span className="text-[11px] uppercase tracking-wide text-slate-400">Fee</span>
                    <div className="flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-1">
                      <span className="text-xs text-slate-500">£</span>
                      <input
                        type="number"
                        step="0.01"
                        className="w-20 border-0 bg-transparent p-0 text-sm focus:outline-none"
                        value={feeInputValue}
                        onClick={event => event.stopPropagation()}
                        onChange={event => handleSingerFeeChange(singer, event.target.value)}
                      />
                    </div>
                  </div>
                ) : null}
              </div>
            );
          })}
        </div>
        <div className="text-xs text-slate-400">Click a singer to toggle them, then adjust the fee if needed.</div>
      </div>
    );
  };

  useEffect(() => {
    const base = totals.base || 0;
    const custom = Number(formState.pricing_custom_fees) || 0;
    const discount = Number(formState.pricing_discount) || 0;
    const total = Math.max(base + custom - discount, 0);
    const nextTotal = total.toFixed(2);
    const currentTotal = formState.pricing_total ? Number(formState.pricing_total).toFixed(2) : '';
    if (currentTotal !== nextTotal) {
      onChange('pricing_total', nextTotal);
    }
  }, [totals.base, formState.pricing_custom_fees, formState.pricing_discount, formState.pricing_total, onChange]);

  useEffect(() => {
    if (!formState.pricing_total) return;
    const total = Number(formState.pricing_total);
    if (Number.isFinite(total) && total > 0) {
      onChange('ahmen_fee', total.toFixed(2));
    }
  }, [formState.pricing_total, onChange]);

  const handleSelectService = (serviceId) => {
    const current = formState.pricing_service_id != null ? String(formState.pricing_service_id) : '';
    const target = serviceId != null ? String(serviceId) : '';
    const nextValue = target === current ? '' : target;
    onChange('pricing_service_id', nextValue);
  };

  return (
    <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-4">
      <div>
        <div className="flex items-center justify-between">
          <span className="text-sm font-medium text-slate-600">Service configuration</span>
          {selectedService ? (
            <span className="text-xs text-slate-400">{selectedService.singers.length} preset singers</span>
          ) : null}
        </div>
        <div className="mt-2 flex flex-wrap gap-2">
          {serviceTypes.length ? serviceTypes.map(type => {
            const typeId = type.id != null ? String(type.id) : '';
            const isActive = typeId === (formState.pricing_service_id != null ? String(formState.pricing_service_id) : '');
            return (
              <button
                key={type.id}
                type="button"
                onClick={() => handleSelectService(type.id)}
                className={`inline-flex items-center gap-2 rounded-full border px-3 py-1.5 text-xs font-medium transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${isActive ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:border-indigo-200 hover:text-indigo-600'}`}
              >
                {type.label}
              </button>
            );
          }) : (
            <span className="text-sm text-slate-500">No service templates configured.</span>
          )}
        </div>
      </div>

      {renderSingers()}

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
  formState,
  onChange,
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
    onChange(prev => {
      const next = applyDerivedFields({ ...prev, [name]: value });
      return next;
    });
  };

  const [savedVenueId, setSavedVenueId] = useState(() => (
    formState.venue_id ? String(formState.venue_id) : ''
  ));

  useEffect(() => {
    setSavedVenueId(formState.venue_id ? String(formState.venue_id) : '');
  }, [formState.venue_id]);

  const [showVenueModal, setShowVenueModal] = useState(false);
  const [venueDraft, setVenueDraft] = useState({
    name: '',
    address1: '',
    address2: '',
    address3: '',
    town: '',
    postcode: '',
    is_private: false
  });

  const openVenueModal = () => {
    setVenueDraft({
      name: formState.venue_name || '',
      address1: formState.venue_address1 || '',
      address2: formState.venue_address2 || '',
      address3: formState.venue_address3 || '',
      town: formState.venue_town || '',
      postcode: formState.venue_postcode || '',
      is_private: Boolean(formState.venue_same_as_client)
    });
    setShowVenueModal(true);
  };

  const closeVenueModal = () => {
    setShowVenueModal(false);
  };

  const handleVenueDraftChange = (field, value) => {
    setVenueDraft(prev => ({ ...prev, [field]: value }));
  };

  const handleCreateVenue = async () => {
    if (!venueDraft.name.trim()) return;
    const result = await onSaveVenue({ ...venueDraft });
    if (result) {
      setShowVenueModal(false);
    }
  };

  const [activeGroupKey, setActiveGroupKey] = useState(() => {
    const defaultGroup = FORM_GROUPS.find(group => group.defaultOpen) || FORM_GROUPS[0];
    return defaultGroup ? defaultGroup.key : null;
  });

  const activeGroup = useMemo(() => (
    FORM_GROUPS.find(group => group.key === activeGroupKey) || FORM_GROUPS[0] || null
  ), [activeGroupKey]);

  useEffect(() => {
    if (!activeGroup && FORM_GROUPS.length) {
      setActiveGroupKey(FORM_GROUPS[0].key);
    }
  }, [activeGroup]);

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

  return (
    <>
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

      <div className="flex flex-col gap-6 lg:flex-row">
        <nav className="lg:w-64 flex-shrink-0">
          <div className="space-y-2" role="tablist" aria-orientation="vertical">
            {FORM_GROUPS.map(group => {
              const isActive = activeGroup?.key === group.key;
              return (
                <button
                  key={group.key}
                  type="button"
                  role="tab"
                  aria-selected={isActive}
                  onClick={() => setActiveGroupKey(group.key)}
                  className={`w-full text-left rounded-lg border px-4 py-3 transition focus:outline-none focus:ring-2 focus:ring-indigo-500 ${isActive ? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold shadow-sm' : 'border-transparent bg-white text-slate-600 hover:bg-slate-50 hover:border-slate-200'}`}
                >
                  <div className="text-sm font-semibold">{group.title}</div>
                  {group.description ? (
                    <p className="mt-1 text-xs text-slate-500">{group.description}</p>
                  ) : null}
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
                    return (
                      <PricingPanel
                        key={field.name}
                        pricingConfig={pricingConfig}
                        pricingTotals={pricingTotals}
                        formState={formState}
                        onChange={handleFieldChange}
                        hasExisting={hasExisting}
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
                        onCreateNew={openVenueModal}
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
                      readOnly={field.readOnly}
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
              <label className="block text-sm font-medium text-slate-600">
                Venue name
                <input
                  type="text"
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  value={venueDraft.name}
                  onChange={event => handleVenueDraftChange('name', event.target.value)}
                />
              </label>
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
              <label className="flex items-center gap-2 text-sm font-medium text-slate-600">
                <input
                  type="checkbox"
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                  checked={venueDraft.is_private}
                  onChange={event => handleVenueDraftChange('is_private', event.target.checked)}
                />
                Private residence (use client address)
              </label>
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

function BusinessWorkspace({ business, onSwitch }) {
  const [jobsheets, setJobsheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [listLoading, setListLoading] = useState(true);
  const [sortConfig, setSortConfig] = useState({ key: 'event_date', direction: 'desc' });
  const [deletingId, setDeletingId] = useState(null);
  const [statusUpdatingId, setStatusUpdatingId] = useState(null);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');

  const normalizeJobsheet = useCallback(item => ({
    ...item,
    status: normalizeStatus(item.status) || 'enquiry'
  }), []);

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
      setJobsheets((data || []).map(normalizeJobsheet));
    } catch (err) {
      console.error('Failed to refresh jobsheets', err);
      setError(err?.message || 'Unable to refresh jobsheets');
    } finally {
      setListLoading(false);
    }
  }, [business.id, normalizeJobsheet]);

  useEffect(() => {
    let mounted = true;
    (async () => {
      setLoading(true);
      setError('');
      await refreshJobsheets();
      if (mounted) setLoading(false);
    })();
    return () => { mounted = false; };
  }, [refreshJobsheets]);

  useEffect(() => {
    if (!window.api || typeof window.api.onJobsheetChange !== 'function') return () => {};
    const unsubscribe = window.api.onJobsheetChange(payload => {
      if (!payload || payload.businessId !== business.id) return;
      if (payload.type === 'jobsheet-updated' && payload.snapshot) {
        mergeJobsheetSnapshot(payload.snapshot);
      } else {
        refreshJobsheets();
      }
    });
    return () => unsubscribe?.();
  }, [business.id, refreshJobsheets, mergeJobsheetSnapshot]);

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
    openJobsheetWindow(undefined);
  }, [openJobsheetWindow]);

  const handleOpenExisting = useCallback((jobsheetId) => {
    if (!jobsheetId) return;
    openJobsheetWindow(jobsheetId);
  }, [openJobsheetWindow]);

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

      <main className="max-w-7xl mx-auto px-6 py-6 space-y-4">
        {error ? <div className="rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
        {message ? <div className="rounded border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-700">{message}</div> : null}
        {loading ? (
          <div className="bg-white rounded-lg border border-slate-200 p-6 text-center text-slate-500">Loading workspace…</div>
        ) : (
          <>
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
            />
            <div className="rounded-lg border border-dashed border-slate-300 bg-white px-6 py-8 text-sm text-slate-500">
              Jobsheets open in a dedicated window. Changes save automatically and this list refreshes when the editor window makes updates.
            </div>
          </>
        )}
      </main>
    </div>
  );
}

function JobsheetEditorWindow({ businessId, businessName, initialJobsheetId }) {
  const numericBusinessId = Number(businessId) || 0;
  const [business, setBusiness] = useState(businessName ? { id: numericBusinessId, business_name: businessName } : null);
  const [formState, setFormState] = useState(DEFAULT_JOBSHEET(numericBusinessId));
  const [jobsheetId, setJobsheetId] = useState(initialJobsheetId && initialJobsheetId !== 'new' ? Number(initialJobsheetId) : null);
  const [venues, setVenues] = useState([]);
  const [pricingConfig, setPricingConfig] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [venueSaving, setVenueSaving] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const formStateRef = useRef(DEFAULT_JOBSHEET(numericBusinessId));

  const autoSaveTimer = useRef(null);
  const initialLoadRef = useRef(true);
  const creatingRef = useRef(false);

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
        const [businessList, venueData, pricingData] = await Promise.all([
          api.businessSettings(),
          api.getAhmenVenues({ businessId: numericBusinessId }),
          api.getAhmenPricing()
        ]);
        if (!mounted) return;
        const businessRecord = business || (businessList || []).find(item => item.id === numericBusinessId) || null;
        setBusiness(businessRecord);
        setVenues(normalizeVenues(venueData));
        setPricingConfig(pricingData || null);

        let effectiveJobsheetId = jobsheetId;
        if (!effectiveJobsheetId && initialJobsheetId && initialJobsheetId !== 'new') {
          effectiveJobsheetId = Number(initialJobsheetId);
          setJobsheetId(effectiveJobsheetId);
        }

        if (effectiveJobsheetId) {
          const sheet = await api.getAhmenJobsheet(effectiveJobsheetId);
          if (!mounted) return;
          if (sheet) {
            setFormState(mapApiToForm(sheet, numericBusinessId));
          }
        } else {
          setFormState(DEFAULT_JOBSHEET(numericBusinessId));
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
  }, [numericBusinessId, business, initialJobsheetId, jobsheetId]);

  useEffect(() => {
    formStateRef.current = formState;
  }, [formState]);

  const pricingDerived = useMemo(() => {
    if (!pricingConfig) return null;
    const service = pricingConfig.serviceTypes?.find(type => type.id === formState.pricing_service_id);
    const selectedEntries = normalizeSingerEntries(formState.pricing_selected_singers);
    let base = 0;
    let singerCount = 0;
    if (service) {
      const serviceSingerList = Array.isArray(service.singers) ? service.singers : [];
      const serviceSingerMap = new Map(serviceSingerList.map(singer => [String(singer.id), singer]));
      selectedEntries.forEach(entry => {
        const singer = serviceSingerMap.get(entry.id);
        if (!singer) return;
        const feeValue = entry.fee !== undefined && entry.fee !== null && entry.fee !== ''
          ? Number(entry.fee)
          : Number(singer.fee);
        base += Number.isFinite(feeValue) ? feeValue : 0;
        singerCount += 1;
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
  }, [pricingConfig, formState]);

  useEffect(() => {
    if (!pricingDerived) return;
    setFormState(prev => {
      const nextTotal = pricingDerived.totalString || '';
      const currentTotal = prev.pricing_total ?? '';
      const shouldUpdateTotal = nextTotal !== currentTotal;

      let shouldUpdateFee = false;
      let nextFeeValue = prev.ahmen_fee ?? '';
      if (pricingDerived.hasSelection) {
        const candidateFee = pricingDerived.totalString || '';
        if (candidateFee && candidateFee !== (prev.ahmen_fee ?? '')) {
          shouldUpdateFee = true;
          nextFeeValue = candidateFee;
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
        const url = new URL(window.location.href);
        url.searchParams.set('jobsheetId', newId);
        window.history.replaceState({}, '', url.toString());
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
  }, [loading, jobsheetId, numericBusinessId, formState]);

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
      const updatedVenues = await api.getAhmenVenues({ businessId: numericBusinessId });
      const normalized = normalizeVenues(updatedVenues);
      setVenues(normalized);

      const savedVenueId = result?.venue_id ?? payload.venue_id ?? null;
      if (savedVenueId) {
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

  const handleDelete = useCallback(async () => {
    if (!jobsheetId) {
      window.close();
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
      window.close();
    } catch (err) {
      console.error('Failed to delete jobsheet', err);
      setError(err?.message || 'Unable to delete jobsheet');
    } finally {
      setSaving(false);
    }
  }, [jobsheetId, numericBusinessId]);

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

  const resolvedBusiness = business || { id: numericBusinessId, business_name: businessName || 'Jobsheet' };

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
        {loading ? (
          <div className="bg-white rounded-lg border border-slate-200 p-6 text-center text-slate-500">Loading jobsheet…</div>
        ) : (
          <>
            <div className="sticky top-0 z-20 -mx-6 px-6 pt-2 pb-4 bg-slate-100/95 backdrop-blur">
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
                  <div className="text-base font-semibold text-slate-800">{toCurrency(formState.ahmen_fee || pricingDerived?.total || 0)}</div>
                </div>
                <div>
                  <div className="text-xs uppercase tracking-wide text-slate-400">Status</div>
                  <span className={`inline-flex items-center rounded-full px-3 py-1 text-xs font-semibold ${STATUS_STYLES[formState.status] || STATUS_STYLES.enquiry}`}>
                    {STATUS_OPTIONS.find(opt => opt.value === formState.status)?.label || 'Enquiry'}
                  </span>
                </div>
              </div>
            </div>

            <JobsheetEditor
              business={resolvedBusiness}
              formState={formState}
              onChange={setFormState}
              onDelete={handleDelete}
              saving={saving}
              deleting={false}
              hasExisting={Boolean(jobsheetId)}
              venues={venues}
              onSaveVenue={handleSaveVenue}
              venueSaving={venueSaving}
              pricingConfig={pricingConfig}
              pricingTotals={pricingDerived}
            />
          </>
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

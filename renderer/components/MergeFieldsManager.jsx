import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';

const EMPTY_FIELD = {
  field_key: '',
  label: '',
  placeholder: '',
  category: '',
  description: '',
  show_in_jobsheet: true,
  active: true,
  bindings: [
    {
      template: 'ahmen_excel',
      sheet: 'Client Data',
      cell: '',
      data_type: 'string',
      format: ''
    }
  ]
};

const DATA_TYPES = [
  { value: 'string', label: 'String' },
  { value: 'number', label: 'Number' }
];

const STORAGE_KEY = 'mergeFieldManager:window';

const DEFAULT_FIELD_MAPPING_LABELS = {
  client_name: 'Client name (client record or jobsheet)',
  client_email: 'Client email (client record or jobsheet)',
  client_phone: 'Client phone (client record or jobsheet)',
  client_address1: 'Client address line 1 (client record or jobsheet)',
  client_address2: 'Client address line 2 (client record or jobsheet)',
  client_address3: 'Client address line 3 (client record or jobsheet)',
  client_town: 'Client town (client record or jobsheet)',
  client_postcode: 'Client postcode (client record or jobsheet)',
  event_type: 'Event type (event record or jobsheet)',
  event_date: 'Event date (event record or jobsheet)',
  event_start: 'Event start time (event record or jobsheet)',
  event_end: 'Event end time (event record or jobsheet)',
  venue_name: 'Venue name (event record or jobsheet)',
  venue_address1: 'Venue address line 1 (event record or jobsheet)',
  venue_address2: 'Venue address line 2 (event record or jobsheet)',
  venue_address3: 'Venue address line 3 (event record or jobsheet)',
  venue_town: 'Venue town (event record or jobsheet)',
  venue_postcode: 'Venue postcode (event record or jobsheet)',
  caterer_name: 'Caterer name (event record or jobsheet)',
  ahmen_fee: 'AhMen fee (jobsheet or pricing)',
  total_amount: 'Total fees (AhMen fee + extra fees + production fees)',
  extra_fees: 'Extra fees (jobsheet custom fees or calculated)',
  production_fees: 'Production fees (jobsheet or calculated)',
  deposit_amount: 'Deposit amount (jobsheet or calculated)',
  balance_amount: 'Balance amount (jobsheet or calculated)',
  balance_due_date: 'Balance due date (jobsheet or calculated)',
  balance_reminder_date: 'Balance reminder date (jobsheet or calculated)',
  service_types: 'Service types (jobsheet)',
  specialist_singers: 'Specialist singers (jobsheet)'
};

const SOURCE_OPTION_GROUPS = [
  {
    label: 'Document totals',
    options: [
      { value: 'context.totalAmount', label: 'Document total (calculated)' },
      { value: 'context.extraFees', label: 'Extra fees (calculated)' },
      { value: 'context.productionFees', label: 'Production fees (calculated)' },
      { value: 'context.depositAmount', label: 'Deposit amount (calculated)' },
      { value: 'context.balanceAmount', label: 'Balance amount (calculated)' },
      { value: 'context.balanceDate', label: 'Balance due date (calculated)' },
      { value: 'context.balanceRemind', label: 'Balance reminder date (calculated)' }
    ]
  },
  {
    label: 'Jobsheet fields',
    options: [
      { value: 'jobsheet.client_name', label: 'Client name' },
      { value: 'jobsheet.client_email', label: 'Client email' },
      { value: 'jobsheet.client_phone', label: 'Client phone' },
      { value: 'jobsheet.client_address1', label: 'Client address 1' },
      { value: 'jobsheet.client_address2', label: 'Client address 2' },
      { value: 'jobsheet.client_address3', label: 'Client address 3' },
      { value: 'jobsheet.client_town', label: 'Client town' },
      { value: 'jobsheet.client_postcode', label: 'Client postcode' },
      { value: 'jobsheet.event_type', label: 'Event type' },
      { value: 'jobsheet.event_date', label: 'Event date' },
      { value: 'jobsheet.event_start', label: 'Event start' },
      { value: 'jobsheet.event_end', label: 'Event end' },
      { value: 'jobsheet.venue_name', label: 'Venue name' },
      { value: 'jobsheet.venue_address1', label: 'Venue address 1' },
      { value: 'jobsheet.venue_address2', label: 'Venue address 2' },
      { value: 'jobsheet.venue_address3', label: 'Venue address 3' },
      { value: 'jobsheet.venue_town', label: 'Venue town' },
      { value: 'jobsheet.venue_postcode', label: 'Venue postcode' },
      { value: 'jobsheet.caterer_name', label: 'Caterer name' },
      { value: 'jobsheet.ahmen_fee', label: 'AhMen fee' },
      { value: 'jobsheet.pricing_total', label: 'Pricing total' },
      { value: 'jobsheet.extra_fees', label: 'Extra fees (jobsheet)' },
      { value: 'jobsheet.pricing_custom_fees', label: 'Custom fees' },
      { value: 'jobsheet.pricing_production_total', label: 'Production total' },
      { value: 'jobsheet.production_fees', label: 'Production fees (jobsheet)' },
      { value: 'jobsheet.deposit_amount', label: 'Deposit amount' },
      { value: 'jobsheet.balance_amount', label: 'Balance amount' },
      { value: 'jobsheet.balance_due_date', label: 'Balance due date' },
      { value: 'jobsheet.balance_reminder_date', label: 'Balance reminder date' },
      { value: 'jobsheet.service_types', label: 'Service types' },
      { value: 'jobsheet.specialist_singers', label: 'Specialist singers' }
    ]
  },
  {
    label: 'Client record',
    options: [
      { value: 'client.name', label: 'Client name' },
      { value: 'client.email', label: 'Client email' },
      { value: 'client.phone', label: 'Client phone' },
      { value: 'client.address1', label: 'Client address 1' },
      { value: 'client.address2', label: 'Client address 2' },
      { value: 'client.address3', label: 'Client address 3' },
      { value: 'client.town', label: 'Client town' },
      { value: 'client.postcode', label: 'Client postcode' }
    ]
  },
  {
    label: 'Event record',
    options: [
      { value: 'event.event_name', label: 'Event name' },
      { value: 'event.type', label: 'Event type' },
      { value: 'event.event_date', label: 'Event date' },
      { value: 'event.startTime', label: 'Event start' },
      { value: 'event.endTime', label: 'Event end' },
      { value: 'event.venue_name', label: 'Venue name' },
      { value: 'event.venue_address1', label: 'Venue address 1' },
      { value: 'event.venue_address2', label: 'Venue address 2' },
      { value: 'event.venue_address3', label: 'Venue address 3' },
      { value: 'event.town', label: 'Event town' },
      { value: 'event.postcode', label: 'Event postcode' },
      { value: 'event.caterer_name', label: 'Caterer name' }
    ]
  },
  {
    label: 'Pricing summary',
    options: [
      { value: 'pricing.total', label: 'Pricing total' },
      { value: 'pricing.extraFees', label: 'Pricing extra fees' },
      { value: 'pricing.productionTotal', label: 'Pricing production total' },
      { value: 'pricing.deposit', label: 'Pricing deposit' },
      { value: 'pricing.balance', label: 'Pricing balance' }
    ]
  },
  {
    label: 'Business record',
    options: [
      { value: 'business.business_name', label: 'Business name' }
    ]
  }
];

function MergeFieldsManager({ onClose, inline = false }) {
  const isInline = Boolean(inline);
  const [fields, setFields] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);
  const [editingField, setEditingField] = useState(null);
  const [formState, setFormState] = useState(EMPTY_FIELD);
  const [confirmDeleteKey, setConfirmDeleteKey] = useState('');
  const [valueSources, setValueSources] = useState({});
  const [sourceSavingKey, setSourceSavingKey] = useState('');
  const [mappingError, setMappingError] = useState('');
  const initialWindowState = useMemo(() => {
    const fallback = {
      width: 860,
      height: 520,
      top: 80,
      left: 80
    };

    if (isInline) return fallback;
    if (typeof window === 'undefined') return fallback;

    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    let stored = null;
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch (err) {
      console.warn('Unable to read merge field manager window state', err);
    }

    const defaultWidth = Math.min(viewportWidth - 120, 960);
    const defaultHeight = Math.min(viewportHeight - 120, 600);

    const width = stored?.width && Number.isFinite(stored.width)
      ? Math.min(Math.max(560, stored.width), viewportWidth - 80)
      : Math.max(640, defaultWidth);

    const height = stored?.height && Number.isFinite(stored.height)
      ? Math.min(Math.max(360, stored.height), viewportHeight - 80)
      : Math.max(420, defaultHeight);

    const left = stored?.left && Number.isFinite(stored.left)
      ? Math.min(Math.max(20, stored.left), Math.max(20, viewportWidth - width - 40))
      : Math.max(40, (viewportWidth - width) / 2);

    const top = stored?.top && Number.isFinite(stored.top)
      ? Math.min(Math.max(20, stored.top), Math.max(20, viewportHeight - height - 40))
      : Math.max(40, (viewportHeight - height) / 2);

    return { width, height, top, left };
  }, [isInline]);

  const [position, setPosition] = useState({ top: initialWindowState.top, left: initialWindowState.left });
  const [size, setSize] = useState({ width: initialWindowState.width, height: initialWindowState.height });
  const dragOffsetRef = useRef({ x: 0, y: 0 });
  const resizeStartRef = useRef({ x: 0, y: 0, width: initialWindowState.width, height: initialWindowState.height });
  const positionRef = useRef(position);
  const sizeRef = useRef(size);
  const [dragging, setDragging] = useState(false);
  const [resizing, setResizing] = useState(false);
  const overlayRef = useRef(null);

  const sourceOptionLookup = useMemo(() => {
    const map = {};
    SOURCE_OPTION_GROUPS.forEach(group => {
      group.options.forEach(option => {
        map[option.value] = option.label;
      });
    });
    return map;
  }, []);

  const refreshValueSources = useCallback(async (fieldKeys) => {
    if (!Array.isArray(fieldKeys) || !fieldKeys.length) {
      setValueSources({});
      return;
    }
    const api = window.api;
    if (!api || typeof api.getMergeFieldValueSources !== 'function') {
      return;
    }
    try {
      const response = await api.getMergeFieldValueSources(fieldKeys);
      setValueSources(response || {});
      setMappingError('');
    } catch (err) {
      console.error('Failed to load placeholder mappings', err);
      setMappingError(err?.message || 'Unable to load placeholder mappings');
    }
  }, []);

  const loadFields = useCallback(async () => {
    try {
      setLoading(true);
      const api = window.api;
      if (!api || typeof api.getMergeFields !== 'function') {
        throw new Error('Placeholder API unavailable');
      }
      const list = await api.getMergeFields();
      const normalized = Array.isArray(list) ? list : [];
      setFields(normalized);
      await refreshValueSources(normalized.map(field => field.field_key));
      setError('');
    } catch (err) {
      console.error('Failed to load merge fields', err);
      setError(err?.message || 'Unable to load placeholders');
    } finally {
      setLoading(false);
    }
  }, [refreshValueSources]);

  useEffect(() => {
    loadFields();
  }, [loadFields]);

  useEffect(() => {
    positionRef.current = position;
  }, [position]);

  useEffect(() => {
    sizeRef.current = size;
  }, [size]);

  const persistWindowState = useCallback(() => {
    if (isInline || typeof window === 'undefined') return;
    try {
      const payload = {
        width: sizeRef.current.width,
        height: sizeRef.current.height,
        top: positionRef.current.top,
        left: positionRef.current.left
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (err) {
      console.warn('Unable to persist merge field manager window state', err);
    }
  }, [isInline]);

  const handleClose = useCallback(() => {
    persistWindowState();
    onClose?.();
  }, [persistWindowState, onClose]);

  useEffect(() => {
    if (isInline || !dragging) return undefined;

    const handleMove = (event) => {
      const offset = dragOffsetRef.current;
      const nextLeft = event.clientX - offset.x;
      const nextTop = event.clientY - offset.y;
      setPosition({ left: nextLeft, top: nextTop });
    };

    const handleUp = () => {
      setDragging(false);
      persistWindowState();
    };

    window.addEventListener('mousemove', handleMove);
    window.addEventListener('mouseup', handleUp);

    return () => {
      window.removeEventListener('mousemove', handleMove);
      window.removeEventListener('mouseup', handleUp);
    };
  }, [dragging, persistWindowState, isInline]);

  useEffect(() => {
    if (isInline || !resizing) return undefined;

    const handleMove = (event) => {
      const start = resizeStartRef.current;
      const deltaX = event.clientX - start.x;
      const deltaY = event.clientY - start.y;
      const nextWidth = Math.max(560, start.width + deltaX);
      const nextHeight = Math.max(360, start.height + deltaY);
      setSize({ width: nextWidth, height: nextHeight });
    };

    const handleUp = () => {
      setResizing(false);
      persistWindowState();
    };

    window.addEventListener('mousemove', handleMove);
    window.addEventListener('mouseup', handleUp);

    return () => {
      window.removeEventListener('mousemove', handleMove);
      window.removeEventListener('mouseup', handleUp);
    };
  }, [resizing, persistWindowState, isInline]);

  const resetForm = () => {
    setFormState(EMPTY_FIELD);
    setEditingField(null);
  };

  const handleEdit = (field) => {
    setEditingField(field.field_key);
    setFormState({
      field_key: field.field_key,
      label: field.label,
      placeholder: field.placeholder || '',
      category: field.category || '',
      description: field.description || '',
      show_in_jobsheet: field.show_in_jobsheet ?? true,
      active: field.active ?? true,
      bindings: (field.bindings && field.bindings.length)
        ? field.bindings.map(binding => ({
            template: binding.template || 'ahmen_excel',
            sheet: binding.sheet || '',
            cell: binding.cell || '',
            data_type: binding.data_type || 'string',
            format: binding.format || '',
            style: binding.style || null
          }))
        : [{ template: 'ahmen_excel', sheet: 'Client Data', cell: '', data_type: 'string', format: '' }]
    });
  };

  const handleAddBinding = () => {
    setFormState(prev => ({
      ...prev,
      bindings: [...prev.bindings, { template: 'ahmen_excel', sheet: '', cell: '', data_type: 'string', format: '' }]
    }));
  };

  const handleRemoveBinding = (index) => {
    setFormState(prev => ({
      ...prev,
      bindings: prev.bindings.filter((_, i) => i !== index)
    }));
  };

  const handleBindingChange = (index, key, value) => {
    setFormState(prev => ({
      ...prev,
      bindings: prev.bindings.map((binding, i) => {
        if (i !== index) return binding;
        return {
          ...binding,
          [key]: value
        };
      })
    }));
  };

  const handleSelectSource = useCallback(async (fieldKey, selectedValue) => {
    if (!fieldKey || !selectedValue) return;
    const api = window.api;
    if (!api) return;

    try {
      setSourceSavingKey(fieldKey);
      setMappingError('');

      if (selectedValue === '__DEFAULT__') {
        if (typeof api.clearMergeFieldValueSource === 'function') {
          await api.clearMergeFieldValueSource(fieldKey);
        }
        setValueSources(prev => {
          const next = { ...prev };
          delete next[fieldKey];
          return next;
        });
        return;
      }

      let nextPath = selectedValue;
      if (selectedValue === '__CUSTOM__') {
        const existingPath = valueSources[fieldKey]?.source_path || '';
        const input = window.prompt('Enter the context path to use (e.g. jobsheet.ahmen_fee):', existingPath);
        if (!input) return;
        nextPath = input.trim();
      }

      if (!nextPath || nextPath === '__DEFAULT__') {
        return;
      }

      if (!nextPath.includes('.')) {
        setMappingError('Enter a full context path such as jobsheet.ahmen_fee');
        return;
      }

      if (typeof api.setMergeFieldValueSource !== 'function') return;

      await api.setMergeFieldValueSource(fieldKey, {
        source_type: 'contextPath',
        source_path: nextPath
      });

      setValueSources(prev => ({
        ...prev,
        [fieldKey]: {
          field_key: fieldKey,
          source_type: 'contextPath',
          source_path: nextPath,
          literal_value: null
        }
      }));
    } catch (err) {
      console.error('Failed to update placeholder mapping', err);
      setMappingError(err?.message || 'Unable to update placeholder mapping');
    } finally {
      setSourceSavingKey('');
    }
  }, [valueSources]);

  const validateForm = () => {
    if (!formState.field_key || !formState.field_key.trim()) {
      return 'Field key is required';
    }
    if (!formState.label || !formState.label.trim()) {
      return 'Label is required';
    }
    if (!formState.bindings.length) {
      return 'At least one binding is required';
    }
    const invalidBinding = formState.bindings.find(binding => !binding.template);
    if (invalidBinding) {
      return 'Each binding must specify a template';
    }
    return null;
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    const validationError = validateForm();
    if (validationError) {
      setError(validationError);
      return;
    }

    try {
      setSaving(true);
      setError('');
      const api = window.api;
      if (!api || typeof api.saveMergeField !== 'function') {
        throw new Error('Placeholder API unavailable');
      }
      const payload = {
        ...formState,
        bindings: formState.bindings.filter(binding => binding.template)
      };
      await api.saveMergeField(payload);
      await loadFields();
      resetForm();
    } catch (err) {
      console.error('Failed to save merge field', err);
      setError(err?.message || 'Unable to save placeholder');
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (fieldKey) => {
    if (!fieldKey) return;
    try {
      const api = window.api;
      if (!api || typeof api.deleteMergeField !== 'function') {
        throw new Error('Placeholder API unavailable');
      }
      await api.deleteMergeField(fieldKey);
      await loadFields();
    } catch (err) {
      console.error('Failed to delete merge field', err);
      setError(err?.message || 'Unable to delete placeholder');
    } finally {
      setConfirmDeleteKey('');
    }
  };

  const categories = useMemo(() => {
    const all = new Set(fields.map(field => field.category).filter(Boolean));
    if (formState.category && !all.has(formState.category)) {
      all.add(formState.category);
    }
    return Array.from(all).sort();
  }, [fields, formState.category]);

  const handleDragStart = (event) => {
    if (isInline) return;
    event.preventDefault();
    const bounds = event.currentTarget.getBoundingClientRect();
    dragOffsetRef.current = {
      x: event.clientX - bounds.left,
      y: event.clientY - bounds.top
    };
    setDragging(true);
  };

  const handleResizeStart = (event) => {
    if (isInline) return;
    event.preventDefault();
    resizeStartRef.current = {
      x: event.clientX,
      y: event.clientY,
      width: sizeRef.current.width,
      height: sizeRef.current.height
    };
    setResizing(true);
  };

  const handleOverlayMouseDown = useCallback((event) => {
    if (isInline) return;
    if (event.target === overlayRef.current) {
      handleClose();
    }
  }, [handleClose, isInline]);

  const headerClass = isInline
    ? 'flex items-center justify-between border-b border-slate-200 pb-4'
    : `flex items-center justify-between border-b border-slate-200 px-6 py-4 cursor-move ${dragging ? 'select-none' : ''}`;

  const headerContent = (
    <div
      className={headerClass}
      onMouseDown={isInline ? undefined : handleDragStart}
    >
      <div>
        <h2 className="text-xl font-semibold text-slate-800">Placeholder manager</h2>
        <p className="text-sm text-slate-500">Add or edit merge fields used across templates.</p>
      </div>
      <button
        onClick={handleClose}
        className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50"
      >
        Close
      </button>
    </div>
  );

  const errorClass = isInline
    ? 'rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700'
    : 'mx-6 mt-4 rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700';

  const contentWrapperClass = isInline
    ? 'space-y-6'
    : 'flex-1 overflow-y-auto px-6 py-4 space-y-6';

  const mappingFields = useMemo(() => {
    if (!Array.isArray(fields)) return [];
    return [...fields]
      .filter(field => field.active !== 0)
      .sort((a, b) => {
        const labelA = (a.label || a.field_key || '').toLowerCase();
        const labelB = (b.label || b.field_key || '').toLowerCase();
        if (labelA < labelB) return -1;
        if (labelA > labelB) return 1;
        return 0;
      });
  }, [fields]);

  const contentSections = (
    <div className={contentWrapperClass}>
      <section className="rounded border border-slate-200 bg-white p-4 space-y-3">
        <div>
          <h3 className="text-sm font-semibold text-slate-700">Data mapping</h3>
          <p className="text-xs text-slate-500">Choose which jobsheet, client, or event fields populate each placeholder when documents are generated.</p>
        </div>
        {mappingError ? (
          <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-600">{mappingError}</div>
        ) : null}
        {loading ? (
          <div className="text-sm text-slate-500">Loading placeholders…</div>
        ) : !mappingFields.length ? (
          <div className="text-sm text-slate-500">No placeholders available yet.</div>
        ) : (
          <div className="max-h-72 overflow-auto rounded border border-slate-200">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-100 text-xs uppercase text-slate-500">
                <tr>
                  <th className="px-3 py-2 text-left">Placeholder</th>
                  <th className="px-3 py-2 text-left">Default data</th>
                  <th className="px-3 py-2 text-left">Override</th>
                </tr>
              </thead>
              <tbody>
                {mappingFields.map(field => {
                  const override = valueSources[field.field_key];
                  let selectedValue = '__DEFAULT__';
                  let overrideLabel = '';

                  if (override && override.source_type === 'contextPath' && override.source_path) {
                    selectedValue = override.source_path;
                    overrideLabel = sourceOptionLookup[override.source_path] || override.source_path;
                  } else if (override && override.source_type === 'literal') {
                    overrideLabel = override.literal_value ? `Literal value: ${override.literal_value}` : 'Literal value';
                  }

                  const optionExists = selectedValue === '__DEFAULT__' || Boolean(sourceOptionLookup[selectedValue]);
                  const defaultLabel = DEFAULT_FIELD_MAPPING_LABELS[field.field_key] || 'Calculated automatically from jobsheet, client, or event data';
                  const isSavingSource = sourceSavingKey === field.field_key;

                  return (
                    <tr key={field.field_key} className="border-b border-slate-200 last:border-b-0">
                      <td className="px-3 py-2 align-top">
                        <div className="font-medium text-slate-700">{field.label || field.field_key}</div>
                        <div className="text-xs text-slate-500">{field.placeholder ? `{{${field.placeholder}}}` : field.field_key}</div>
                      </td>
                      <td className="px-3 py-2 align-top text-sm text-slate-600">{defaultLabel}</td>
                      <td className="px-3 py-2 align-top">
                        <div className="flex flex-col gap-1">
                          <select
                            value={selectedValue}
                            onChange={(event) => handleSelectSource(field.field_key, event.target.value)}
                            disabled={isSavingSource}
                            className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 disabled:cursor-not-allowed disabled:opacity-60"
                          >
                            <option value="__DEFAULT__">Default (calculated)</option>
                            {!optionExists && selectedValue !== '__DEFAULT__' ? (
                              <option value={selectedValue}>{overrideLabel || selectedValue}</option>
                            ) : null}
                            {SOURCE_OPTION_GROUPS.map(group => (
                              <optgroup key={group.label} label={group.label}>
                                {group.options.map(option => (
                                  <option key={option.value} value={option.value}>{option.label}</option>
                                ))}
                              </optgroup>
                            ))}
                            <option value="__CUSTOM__">Custom context path…</option>
                          </select>
                          {isSavingSource ? (
                            <span className="text-xs text-slate-500">Saving…</span>
                          ) : overrideLabel ? (
                            <span className="text-xs text-indigo-600">Override: {overrideLabel}</span>
                          ) : null}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>
      <section className="rounded border border-slate-200 bg-slate-50 p-4">
        <h3 className="text-sm font-semibold text-slate-700 mb-3">Existing placeholders</h3>
        {loading ? (
          <div className="text-sm text-slate-500">Loading placeholders…</div>
        ) : !fields.length ? (
          <div className="text-sm text-slate-500">No placeholders found.</div>
        ) : (
          <table className="min-w-full border border-slate-200 text-sm">
            <thead className="bg-slate-100 text-xs uppercase text-slate-500">
              <tr>
                <th className="border border-slate-200 px-3 py-2 text-left">Key</th>
                <th className="border border-slate-200 px-3 py-2 text-left">Label</th>
                <th className="border border-slate-200 px-3 py-2 text-left">Placeholder</th>
                <th className="border border-slate-200 px-3 py-2 text-left">Category</th>
                <th className="border border-slate-200 px-3 py-2 text-left">Bindings</th>
                <th className="border border-slate-200 px-3 py-2 text-right">Actions</th>
              </tr>
            </thead>
            <tbody>
              {fields.map(field => (
                <tr key={field.field_key} className="odd:bg-white even:bg-slate-50">
                  <td className="border border-slate-200 px-3 py-2 font-mono text-xs text-slate-500">{field.field_key}</td>
                  <td className="border border-slate-200 px-3 py-2 text-slate-700">{field.label}</td>
                  <td className="border border-slate-200 px-3 py-2 text-slate-600">{field.placeholder || '—'}</td>
                  <td className="border border-slate-200 px-3 py-2 text-slate-600">{field.category || '—'}</td>
                  <td className="border border-slate-200 px-3 py-2 text-slate-500">
                    {field.bindings && field.bindings.length ? (
                      <ul className="text-xs space-y-1">
                        {field.bindings.map((binding, index) => (
                          <li key={`${field.field_key}-binding-${index}`}>
                            <span className="font-mono text-slate-600">{binding.template}</span>
                            {binding.sheet && binding.cell ? ` · ${binding.sheet}!${binding.cell}` : ''}
                          </li>
                        ))}
                      </ul>
                    ) : '—'}
                  </td>
                  <td className="border border-slate-200 px-3 py-2 text-right">
                    <div className="flex justify-end gap-2">
                      <button
                        onClick={() => handleEdit(field)}
                        className="rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100"
                      >
                        Edit
                      </button>
                      {confirmDeleteKey === field.field_key ? (
                        <>
                          <button
                            onClick={() => handleDelete(field.field_key)}
                            className="rounded border border-red-200 px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-50"
                          >
                            Confirm
                          </button>
                          <button
                            onClick={() => setConfirmDeleteKey('')}
                            className="rounded border border-slate-200 px-2 py-1 text-xs font-medium text-slate-500 hover:bg-slate-50"
                          >
                            Cancel
                          </button>
                        </>
                      ) : (
                        <button
                          onClick={() => setConfirmDeleteKey(field.field_key)}
                          className="rounded border border-slate-200 px-2 py-1 text-xs font-medium text-slate-400 hover:bg-slate-100"
                        >
                          Delete
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <section className="rounded border border-slate-200 bg-white p-4">
        <h3 className="text-sm font-semibold text-slate-700 mb-3">
          {editingField ? `Edit placeholder: ${editingField}` : 'Add new placeholder'}
        </h3>
        <form className="space-y-4" onSubmit={handleSubmit}>
          <div className="grid gap-4 sm:grid-cols-2">
            <div className="flex flex-col">
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500" htmlFor="merge-field-key">Field key</label>
              <input
                id="merge-field-key"
                type="text"
                value={formState.field_key}
                onChange={event => setFormState(prev => ({ ...prev, field_key: event.target.value }))}
                className="mt-1 rounded border border-slate-300 px-3 py-2 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none"
                placeholder="e.g. client_name"
                readOnly={Boolean(editingField)}
              />
            </div>
            <div className="flex flex-col">
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500" htmlFor="merge-field-label">Label</label>
              <input
                id="merge-field-label"
                type="text"
                value={formState.label}
                onChange={event => setFormState(prev => ({ ...prev, label: event.target.value }))}
                className="mt-1 rounded border border-slate-300 px-3 py-2 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none"
                placeholder="Client Name"
              />
            </div>
            <div className="flex flex-col">
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500" htmlFor="merge-field-placeholder">Placeholder (optional)</label>
              <input
                id="merge-field-placeholder"
                type="text"
                value={formState.placeholder}
                onChange={event => setFormState(prev => ({ ...prev, placeholder: event.target.value }))}
                className="mt-1 rounded border border-slate-300 px-3 py-2 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none"
                placeholder="CLIENT_NAME"
              />
            </div>
            <div className="flex flex-col">
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500" htmlFor="merge-field-category">Category</label>
              <input
                id="merge-field-category"
                list="merge-field-categories"
                value={formState.category}
                onChange={event => setFormState(prev => ({ ...prev, category: event.target.value }))}
                className="mt-1 rounded border border-slate-300 px-3 py-2 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none"
                placeholder="client"
              />
              <datalist id="merge-field-categories">
                {categories.map(category => (
                  <option key={category} value={category} />
                ))}
              </datalist>
            </div>
          </div>

          <div className="flex flex-col">
            <label className="text-xs font-semibold uppercase tracking-wide text-slate-500" htmlFor="merge-field-description">Description</label>
            <textarea
              id="merge-field-description"
              value={formState.description}
              onChange={event => setFormState(prev => ({ ...prev, description: event.target.value }))}
              className="mt-1 rounded border border-slate-300 px-3 py-2 text-sm text-slate-700 focus:border-indigo-500 focus:outline-none"
              rows={2}
              placeholder="Short note about this placeholder"
            />
          </div>

          <div className="grid gap-4 sm:grid-cols-2">
            <label className="inline-flex items-center gap-2 text-sm text-slate-600">
              <input
                type="checkbox"
                checked={!!formState.show_in_jobsheet}
                onChange={event => setFormState(prev => ({ ...prev, show_in_jobsheet: event.target.checked }))}
                className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
              />
              Show on jobsheet editor
            </label>
            <label className="inline-flex items-center gap-2 text-sm text-slate-600">
              <input
                type="checkbox"
                checked={!!formState.active}
                onChange={event => setFormState(prev => ({ ...prev, active: event.target.checked }))}
                className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
              />
              Active
            </label>
          </div>

          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <h4 className="text-sm font-semibold text-slate-700">Template bindings</h4>
              <button
                type="button"
                onClick={handleAddBinding}
                className="inline-flex items-center gap-1 rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
              >
                Add binding
              </button>
            </div>
            {formState.bindings.map((binding, index) => (
              <div key={`binding-${index}`} className="grid gap-2 rounded border border-slate-200 bg-slate-50 p-3 sm:grid-cols-5">
                <div className="flex flex-col">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Template</label>
                  <input
                    type="text"
                    value={binding.template}
                    onChange={event => handleBindingChange(index, 'template', event.target.value)}
                    className="mt-1 rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 focus:border-indigo-500 focus:outline-none"
                  />
                </div>
                <div className="flex flex-col">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Sheet</label>
                  <input
                    type="text"
                    value={binding.sheet || ''}
                    onChange={event => handleBindingChange(index, 'sheet', event.target.value)}
                    className="mt-1 rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 focus:border-indigo-500 focus:outline-none"
                  />
                </div>
                <div className="flex flex-col">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Cell</label>
                  <input
                    type="text"
                    value={binding.cell || ''}
                    onChange={event => handleBindingChange(index, 'cell', event.target.value)}
                    className="mt-1 rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 focus:border-indigo-500 focus:outline-none"
                  />
                </div>
                <div className="flex flex-col">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Data type</label>
                  <select
                    value={binding.data_type || 'string'}
                    onChange={event => handleBindingChange(index, 'data_type', event.target.value)}
                    className="mt-1 rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 focus:border-indigo-500 focus:outline-none"
                  >
                    {DATA_TYPES.map(option => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                </div>
                <div className="flex flex-col">
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Format</label>
                  <input
                    type="text"
                    value={binding.format || ''}
                    onChange={event => handleBindingChange(index, 'format', event.target.value)}
                    className="mt-1 rounded border border-slate-300 px-2 py-1 text-xs text-slate-700 focus:border-indigo-500 focus:outline-none"
                    placeholder="e.g. date_human"
                  />
                </div>
                <div className="sm:col-span-5 flex justify-end">
                  <button
                    type="button"
                    onClick={() => handleRemoveBinding(index)}
                    className="inline-flex items-center gap-1 rounded border border-slate-200 px-2 py-1 text-xs font-medium text-slate-500 hover:bg-slate-100"
                  >
                    Remove
                  </button>
                </div>
              </div>
            ))}
          </div>

          <div className="flex items-center justify-between border-t border-slate-200 pt-4">
            <button
              type="button"
              onClick={resetForm}
              className="inline-flex items-center rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
            >
              Reset
            </button>
            <button
              type="submit"
              disabled={saving}
              className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-semibold text-white hover:bg-indigo-500 disabled:opacity-60"
            >
              {saving ? 'Saving…' : editingField ? 'Update placeholder' : 'Create placeholder'}
            </button>
          </div>
        </form>
      </section>
    </div>
  );

  const body = (
    <>
      {headerContent}
      {error ? <div className={errorClass}>{error}</div> : null}
      {contentSections}
    </>
  );

  if (isInline) {
    return (
      <div className="space-y-4">
        {body}
      </div>
    );
  }

  return (
    <div
      ref={overlayRef}
      onMouseDown={handleOverlayMouseDown}
      className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-slate-900/40 p-6"
    >
      <div
        className="absolute flex max-w-full flex-col overflow-hidden rounded-lg bg-white shadow-xl"
        style={{
          top: position.top,
          left: position.left,
          width: size.width,
          height: size.height,
          minHeight: 420,
          minWidth: 560
        }}
      >
        {body}
        <div
          className="absolute bottom-2 right-2 h-4 w-4 cursor-se-resize rounded-sm border border-slate-300 bg-slate-200"
          onMouseDown={handleResizeStart}
        />
      </div>
    </div>
  );
}

export default MergeFieldsManager;

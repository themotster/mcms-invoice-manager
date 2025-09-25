import React, { useCallback, useEffect, useMemo, useState } from 'react';

const CATEGORY_LABELS = {
  client: 'Client details',
  event: 'Event details',
  venue: 'Venue details',
  financial: 'Financials',
  services: 'Services',
  other: 'Other'
};

const DOC_TYPE_META = {
  invoice: {
    label: 'Invoice',
    filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }]
  },
  quote: {
    label: 'Quote',
    filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }]
  },
  contract: {
    label: 'Contract',
    filters: [{ name: 'Word document', extensions: ['docx'] }]
  },
  workbook: {
    label: 'Jobsheet workbook',
    filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }]
  }
};

function startCase(value) {
  return (value || '')
    .replace(/_/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\w/g, letter => letter.toUpperCase());
}

function slugify(value) {
  return (value || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-+/g, '-');
}

function isWorkbookDefinition(definition) {
  return (definition?.doc_type || '').toLowerCase() === 'workbook';
}

function TemplatesManager({ business, onTemplatesUpdated }) {
  const [placeholders, setPlaceholders] = useState([]);
  const [valueSources, setValueSources] = useState({});
  const [definitions, setDefinitions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadingDefinitions, setLoadingDefinitions] = useState(true);
  const [activeCategory, setActiveCategory] = useState(null);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const [copyFeedback, setCopyFeedback] = useState('');
  const [busyDefinitionKey, setBusyDefinitionKey] = useState('');

  const clearMessageSoon = useCallback(() => {
    if (!message) return;
    const timeout = setTimeout(() => setMessage(''), 2400);
    return () => clearTimeout(timeout);
  }, [message]);

  useEffect(() => clearMessageSoon(), [message, clearMessageSoon]);

  const loadPlaceholders = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const api = window.api;
      if (!api || typeof api.getMergeFields !== 'function') {
        throw new Error('Placeholder API unavailable');
      }
      const list = await api.getMergeFields();
      const normalized = Array.isArray(list) ? list : [];
      setPlaceholders(normalized);

      const keys = normalized.map(field => field.field_key).filter(Boolean);
      if (api.getMergeFieldValueSources && keys.length) {
        try {
          const sources = await api.getMergeFieldValueSources(keys);
          setValueSources(sources || {});
        } catch (sourceErr) {
          console.warn('Unable to load placeholder mappings', sourceErr);
        }
      }
    } catch (err) {
      console.error('Failed to load placeholders', err);
      setError(err?.message || 'Unable to load placeholders');
    } finally {
      setLoading(false);
    }
  }, []);

  const loadDefinitions = useCallback(async () => {
    setLoadingDefinitions(true);
    try {
      const api = window.api;
      if (!api || typeof api.getDocumentDefinitions !== 'function') {
        throw new Error('Template API unavailable');
      }
      const data = await api.getDocumentDefinitions(business.id, { includeInactive: true });
      const allDefinitions = Array.isArray(data) ? data : [];
      const workbookOnly = allDefinitions.filter(isWorkbookDefinition);
      setDefinitions(workbookOnly);
    } catch (err) {
      console.error('Failed to load document definitions', err);
      setError(err?.message || 'Unable to load templates');
    } finally {
      setLoadingDefinitions(false);
    }
  }, [business.id]);

  useEffect(() => {
    loadPlaceholders();
  }, [loadPlaceholders]);

  useEffect(() => {
    loadDefinitions();
  }, [loadDefinitions]);

  const persistDefinition = useCallback(async (definition, overrides = {}) => {
    const api = window.api;
    if (!api || typeof api.saveDocumentDefinition !== 'function') {
      throw new Error('Template API unavailable');
    }

    const sheetExports = overrides.sheet_exports !== undefined
      ? overrides.sheet_exports
      : (definition.sheet_exports || []);

    const payload = {
      key: definition.key,
      doc_type: definition.doc_type,
      label: definition.label,
      description: definition.description,
      file_suffix: definition.file_suffix,
      invoice_variant: definition.invoice_variant,
      template_path: overrides.template_path !== undefined ? overrides.template_path : definition.template_path,
      requires_total: definition.requires_total ? 1 : 0,
      is_primary: definition.is_primary ? 1 : 0,
      is_active: definition.is_active === 0 ? 0 : 1,
      is_locked: definition.is_locked ? 1 : 0,
      sort_order: definition.sort_order,
      sheet_exports: Array.isArray(sheetExports) ? sheetExports.map(entry => ({ ...entry })) : []
    };

    await api.saveDocumentDefinition(business.id, payload);
  }, [business.id]);

  const groupedPlaceholders = useMemo(() => {
    const map = new Map();
    placeholders.forEach(field => {
      const category = (field.category || 'other').toLowerCase();
      if (!map.has(category)) {
        map.set(category, []);
      }
      map.get(category).push(field);
    });

    // ensure deterministic order
    map.forEach(list => {
      list.sort((a, b) => {
        const labelA = (a.label || a.field_key || '').toLowerCase();
        const labelB = (b.label || b.field_key || '').toLowerCase();
        if (labelA < labelB) return -1;
        if (labelA > labelB) return 1;
        return 0;
      });
    });
    return map;
  }, [placeholders]);

  useEffect(() => {
    if (activeCategory && groupedPlaceholders.has(activeCategory)) return;
    const firstCategory = groupedPlaceholders.keys().next().value;
    if (firstCategory) {
      setActiveCategory(firstCategory);
    }
  }, [groupedPlaceholders, activeCategory]);

  const handleCopy = useCallback(async (placeholder) => {
    if (!placeholder) return;
    const token = placeholder.startsWith('{{') ? placeholder : `{{${placeholder}}}`;
    try {
      if (navigator?.clipboard?.writeText) {
        await navigator.clipboard.writeText(token);
      } else {
        const textArea = document.createElement('textarea');
        textArea.value = token;
        textArea.style.position = 'fixed';
        textArea.style.opacity = '0';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }
      setCopyFeedback(`${token} copied`);
      setTimeout(() => setCopyFeedback(''), 2000);
    } catch (err) {
      console.error('Failed to copy placeholder', err);
      setCopyFeedback('Unable to copy to clipboard');
      setTimeout(() => setCopyFeedback(''), 2000);
    }
  }, []);

  const handleReplaceTemplate = useCallback(async (definition) => {
    if (!definition) return;
    const api = window.api;
    if (!api || typeof api.chooseFile !== 'function') {
      setError('File chooser unavailable');
      return;
    }

    const docMeta = DOC_TYPE_META[definition.doc_type?.toLowerCase()] || null;
    try {
      const selectedPath = await api.chooseFile({
        title: `Select template for ${definition.label || startCase(definition.key)}`,
        filters: docMeta?.filters
      });
      if (!selectedPath) return;

      setBusyDefinitionKey(definition.key);
      setMessage('Processing template…');

      if (api.normalizeTemplate) {
        try {
          await api.normalizeTemplate({ templatePath: selectedPath });
        } catch (normalizeErr) {
          console.warn('Normalize template failed', normalizeErr);
        }
      }

      let addedSheets = 0;
      let sheetExportOverrides = null;
      if (isWorkbookDefinition(definition)) {
        try {
          const { sheetExports: nextExports, added } = await buildSheetExportsFromWorkbook(definition, selectedPath);
          addedSheets = added;
          if (added > 0 && Array.isArray(nextExports)) {
            sheetExportOverrides = nextExports;
          }
        } catch (syncErr) {
          console.warn('Unable to sync workbook sheets from template', syncErr);
          setError(syncErr?.message || 'Unable to read workbook sheets');
        }
      }

      const overrides = { template_path: selectedPath };
      if (sheetExportOverrides) {
        overrides.sheet_exports = sheetExportOverrides;
      }

      await persistDefinition(definition, overrides);
      let messageText = `${definition.label || startCase(definition.key)} template updated`;
      if (addedSheets > 0) {
        messageText = `${messageText}. Added PDF exports for ${addedSheets} ${addedSheets === 1 ? 'sheet' : 'sheets'}.`;
      }
      setMessage(messageText);
      await loadDefinitions();
      onTemplatesUpdated?.();
    } catch (err) {
      console.error('Failed to replace template', err);
      setError(err?.message || 'Unable to update template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [buildSheetExportsFromWorkbook, loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleAddSheetExport = useCallback(async (definition) => {
    if (!definition) return;
    if (!isWorkbookDefinition(definition)) {
      setError('Sheet exports can only be added to the workbook template.');
      setTimeout(() => setError(''), 2500);
      return;
    }
    const existingExports = Array.isArray(definition.sheet_exports) ? definition.sheet_exports : [];

    const sheetNameInput = window.prompt('Sheet name to export (case-sensitive):');
    if (!sheetNameInput) return;
    const sheetName = sheetNameInput.trim();
    if (!sheetName) return;

    const defaultLabel = `${definition.label || startCase(definition.key)} – ${sheetName}`;
    const labelInput = window.prompt('Label for this PDF (optional):', defaultLabel);
    if (labelInput === null) return;
    const label = (labelInput || defaultLabel).trim();

    const defaultSuffix = ` - ${sheetName}`;
    const suffixInput = window.prompt('File name suffix (optional):', defaultSuffix);
    if (suffixInput === null) return;
    const fileSuffix = suffixInput || '';

    const newEntry = {
      sheet: sheetName,
      label,
      fileSuffix,
      docType: 'pdf',
      format: 'pdf',
      key: `${definition.key}_${slugify(sheetName)}_pdf`
    };

    const nextExports = existingExports
      .filter(item => (item?.sheet || '').toLowerCase() !== sheetName.toLowerCase())
      .concat([newEntry]);

    try {
      setBusyDefinitionKey(definition.key);
      await persistDefinition(definition, { sheet_exports: nextExports });
      setMessage('Sheet export added');
      await loadDefinitions();
      onTemplatesUpdated?.();
    } catch (err) {
      console.error('Failed to add sheet export', err);
      setError(err?.message || 'Unable to add sheet export');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleRemoveSheetExport = useCallback(async (definition, index) => {
    if (!definition) return;
    if (!isWorkbookDefinition(definition)) {
      setError('Sheet exports can only be managed on the workbook template.');
      setTimeout(() => setError(''), 2500);
      return;
    }
    const existingExports = Array.isArray(definition.sheet_exports) ? definition.sheet_exports : [];
    if (!existingExports[index]) return;

    const entry = existingExports[index];
    const confirmed = window.confirm(`Remove export for sheet "${entry.sheet}"?`);
    if (!confirmed) return;

    const nextExports = existingExports.filter((_, idx) => idx !== index);
    try {
      setBusyDefinitionKey(definition.key);
      await persistDefinition(definition, { sheet_exports: nextExports });
      setMessage('Sheet export removed');
      await loadDefinitions();
      onTemplatesUpdated?.();
    } catch (err) {
      console.error('Failed to remove sheet export', err);
      setError(err?.message || 'Unable to remove sheet export');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const buildSheetExportsFromWorkbook = useCallback(async (definition, templatePath) => {
    if (!isWorkbookDefinition(definition)) {
      return { sheetExports: null, added: 0 };
    }

    const api = window.api;
    if (!api || typeof api.inspectWorkbookSheets !== 'function') {
      throw new Error('Workbook inspection is unavailable');
    }

    const workbookPath = templatePath || definition.template_path;
    if (!workbookPath) {
      throw new Error('Set the workbook template before syncing sheet exports');
    }

    const sheetNames = await api.inspectWorkbookSheets(workbookPath);
    const existing = Array.isArray(definition.sheet_exports) ? definition.sheet_exports.map(entry => ({ ...entry })) : [];
    const seenSheets = new Set(existing.map(entry => (entry?.sheet || '').toLowerCase()));
    const existingKeys = new Set(existing.map(entry => entry?.key).filter(Boolean));
    const additions = [];

    sheetNames.forEach(rawName => {
      const sheetName = (rawName || '').trim();
      if (!sheetName) return;
      const lower = sheetName.toLowerCase();
      if (seenSheets.has(lower)) return;
      const safeDefinitionKey = definition.key || 'workbook';
      const baseSlug = slugify(sheetName) || 'sheet';
      let key = `${safeDefinitionKey}_${baseSlug}_pdf`;
      let counter = 1;
      while (existingKeys.has(key)) {
        counter += 1;
        key = `${safeDefinitionKey}_${baseSlug}_${counter}_pdf`;
      }
      existingKeys.add(key);
      seenSheets.add(lower);
      additions.push({
        sheet: sheetName,
        label: `${definition.label || startCase(definition.key || 'workbook')} – ${sheetName}`,
        fileSuffix: ` - ${sheetName}`,
        docType: 'pdf',
        format: 'pdf',
        key
      });
    });

    if (!additions.length) {
      return { sheetExports: existing, added: 0 };
    }

    return { sheetExports: existing.concat(additions), added: additions.length };
  }, []);

  const handleSyncSheetExports = useCallback(async (definition) => {
    if (!definition) return;
    if (!isWorkbookDefinition(definition)) {
      setError('Only the workbook template can sync sheet exports.');
      setTimeout(() => setError(''), 2500);
      return;
    }

    const workbookPath = definition.template_path;
    if (!workbookPath) {
      setError('Choose a workbook template before syncing sheet exports.');
      setTimeout(() => setError(''), 2500);
      return;
    }

    try {
      setBusyDefinitionKey(definition.key);
      const { sheetExports: nextExports, added } = await buildSheetExportsFromWorkbook(definition, workbookPath);
      if (!Array.isArray(nextExports)) {
        setMessage('No workbook sheets found to export.');
        setTimeout(() => setMessage(''), 2500);
        return;
      }

      if (added === 0) {
        setMessage('All workbook sheets already have PDF exports.');
        setTimeout(() => setMessage(''), 2500);
        return;
      }

      await persistDefinition(definition, { sheet_exports: nextExports });
      setMessage(`Added PDF exports for ${added} ${added === 1 ? 'sheet' : 'sheets'}.`);
      await loadDefinitions();
      onTemplatesUpdated?.();
    } catch (err) {
      console.error('Failed to sync sheet exports', err);
      setError(err?.message || 'Unable to sync sheet exports');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [buildSheetExportsFromWorkbook, loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleNormalizeTemplate = useCallback(async (definition) => {
    if (!definition?.template_path) return;
    if (!isWorkbookDefinition(definition)) {
      setError('Only the workbook template can be normalized here.');
      setTimeout(() => setError(''), 2500);
      return;
    }
    try {
      setBusyDefinitionKey(definition.key);
      await window.api?.normalizeTemplate?.({ templatePath: definition.template_path });
      setMessage('Template normalized');
    } catch (err) {
      console.error('Failed to normalize template', err);
      setError(err?.message || 'Unable to normalize template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, []);

  const handleOpenTemplate = useCallback(async (definition) => {
    if (!definition?.template_path) return;
    try {
      const response = await window.api?.openPath?.(definition.template_path);
      if (response && response.ok === false) {
        throw new Error(response.message || 'Unable to open template');
      }
    } catch (err) {
      console.error('Failed to open template', err);
      setError(err?.message || 'Unable to open template');
    }
  }, []);

  const categoryEntries = useMemo(() => {
    return Array.from(groupedPlaceholders.entries())
      .map(([key, rows]) => ({ key, label: CATEGORY_LABELS[key] || startCase(key), rows }))
      .sort((a, b) => a.label.localeCompare(b.label, 'en', { sensitivity: 'base' }));
  }, [groupedPlaceholders]);

  const activePlaceholders = useMemo(() => {
    if (!activeCategory || !groupedPlaceholders.has(activeCategory)) return [];
    return groupedPlaceholders.get(activeCategory);
  }, [activeCategory, groupedPlaceholders]);

  const activeCategoryLabel = CATEGORY_LABELS[activeCategory] || startCase(activeCategory || '');

  return (
    <div className="space-y-6">
      {error ? <div className="rounded border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div> : null}
      {message ? <div className="rounded border border-indigo-200 bg-indigo-50 px-4 py-3 text-sm text-indigo-700">{message}</div> : null}
      {copyFeedback ? <div className="rounded border border-green-200 bg-green-50 px-4 py-2 text-xs text-green-700">{copyFeedback}</div> : null}

      <div className="flex flex-col gap-6 lg:flex-row">
        <aside className="lg:w-64 flex-shrink-0 space-y-2">
          {loading ? (
            <div className="rounded border border-slate-200 bg-white px-3 py-2 text-sm text-slate-500">Loading placeholders…</div>
          ) : categoryEntries.length === 0 ? (
            <div className="rounded border border-slate-200 bg-white px-3 py-2 text-sm text-slate-500">No placeholders found.</div>
          ) : (
            categoryEntries.map(category => {
              const isActive = activeCategory === category.key;
              return (
                <button
                  key={category.key}
                  type="button"
                  onClick={() => setActiveCategory(category.key)}
                  className={`flex w-full items-center justify-between rounded-lg border px-3 py-2 text-sm transition ${isActive ? 'border-indigo-200 bg-indigo-50 text-indigo-700 font-semibold' : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'}`}
                >
                  <span>{category.label}</span>
                  <span className={`inline-flex h-5 min-w-[1.5rem] items-center justify-center rounded-full text-xs ${isActive ? 'bg-indigo-100 text-indigo-700' : 'bg-slate-100 text-slate-500'}`}>
                    {category.rows.length}
                  </span>
                </button>
              );
            })
          )}
        </aside>

        <section className="flex-1 space-y-4">
          <div className="rounded border border-slate-200 bg-white shadow-sm">
            <header className="border-b border-slate-200 px-4 py-3">
              <h2 className="text-sm font-semibold text-slate-700">Placeholders</h2>
              <p className="mt-1 text-xs text-slate-500">Copy tokens and see where their data comes from.</p>
            </header>
            <div className="px-4 py-3 space-y-2">
              <div className="text-xs uppercase tracking-wide text-slate-500">{activeCategoryLabel}</div>
              <div className="space-y-2">
                {loading ? (
                  <div className="text-sm text-slate-500">Loading…</div>
                ) : !activePlaceholders.length ? (
                  <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-500">No placeholders in this category.</div>
                ) : (
                  activePlaceholders.map(field => {
                    const placeholderToken = field.placeholder || field.field_key;
                    const displayToken = `{{${placeholderToken}}}`;
                    const sourcePath = valueSources[field.field_key]?.source_path || 'Calculated automatically';
                    return (
                      <div key={field.field_key} className="rounded border border-slate-200 px-3 py-2 flex flex-col gap-1 bg-white">
                        <div className="flex items-center justify-between gap-2">
                          <div>
                            <div className="text-sm font-semibold text-slate-700">{field.label || startCase(field.field_key)}</div>
                            <div className="text-xs text-slate-500">{displayToken}</div>
                          </div>
                          <button
                            type="button"
                            onClick={() => handleCopy(placeholderToken)}
                            className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                          >
                            Copy
                          </button>
                        </div>
                        <div className="text-xs text-slate-500">
                          Source: <span className="font-mono text-slate-600">{sourcePath}</span>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          </div>

          <div className="rounded border border-slate-200 bg-white shadow-sm">
            <header className="border-b border-slate-200 px-4 py-3">
              <h2 className="text-sm font-semibold text-slate-700">Templates</h2>
              <p className="mt-1 text-xs text-slate-500">Only the workbook needs a template file. PDF documents are exported from its sheets.</p>
            </header>
            <div className="px-4 py-3 space-y-3">
              {loadingDefinitions ? (
                <div className="text-sm text-slate-500">Loading templates…</div>
              ) : !definitions.length ? (
                <div className="text-sm text-slate-500">No document types configured yet.</div>
              ) : (
                <div className="space-y-3">
                  {definitions.map(definition => {
                    const docMeta = DOC_TYPE_META[definition.doc_type?.toLowerCase()] || null;
                    const workbookDefinition = isWorkbookDefinition(definition);
                    const sheetExports = workbookDefinition && Array.isArray(definition.sheet_exports) ? definition.sheet_exports : [];
                    const templateSummary = workbookDefinition
                      ? (definition.template_path || 'No template selected yet.')
                      : 'Uses the workbook template – no separate file needed.';
                    return (
                      <div key={definition.key} className="rounded border border-slate-200 px-3 py-3">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <div>
                            <div className="text-sm font-semibold text-slate-700">{definition.label || startCase(definition.key)}</div>
                            <div className="text-xs text-slate-500">{docMeta?.label || startCase(definition.doc_type)}</div>
                            <div className="mt-1 text-xs text-slate-400 break-all">{templateSummary}</div>
                          </div>
                          <div className="flex flex-wrap items-center gap-2">
                            {workbookDefinition ? (
                              <>
                                <button
                                  type="button"
                                  disabled={busyDefinitionKey === definition.key}
                                  onClick={() => handleReplaceTemplate(definition)}
                                  className="inline-flex items-center rounded bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-500 disabled:opacity-60"
                                >
                                  {busyDefinitionKey === definition.key ? 'Updating…' : 'Replace workbook template'}
                                </button>
                                <button
                                  type="button"
                                  disabled={!definition.template_path || busyDefinitionKey === definition.key}
                                  onClick={() => handleNormalizeTemplate(definition)}
                                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                                >
                                  Normalize current
                                </button>
                                <button
                                  type="button"
                                  disabled={!definition.template_path}
                                  onClick={() => handleOpenTemplate(definition)}
                                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                                >
                                  Open file
                                </button>
                                <button
                                  type="button"
                                  disabled={!definition.template_path || busyDefinitionKey === definition.key}
                                  onClick={() => handleSyncSheetExports(definition)}
                                  className="inline-flex items-center rounded border border-indigo-200 px-3 py-1.5 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:opacity-60"
                                >
                                  Sync sheet exports
                                </button>
                                <button
                                  type="button"
                                  disabled={busyDefinitionKey === definition.key}
                                  onClick={() => handleAddSheetExport(definition)}
                                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                                >
                                  Add sheet export
                                </button>
                              </>
                            ) : (
                              <span className="text-xs text-slate-500">This document uses the workbook template and its sheet exports.</span>
                            )}
                          </div>
                        </div>
                        <div className="mt-3 space-y-1">
                          <div className="text-xs uppercase tracking-wide text-slate-500">Sheet exports</div>
                          {workbookDefinition ? (
                            sheetExports.length ? (
                              sheetExports.map((entry, idx) => (
                                <div
                                  key={entry.key || `${definition.key}-${entry.sheet || idx}`}
                                  className="flex items-center justify-between gap-2 rounded border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600"
                                >
                                  <div>
                                    <div className="font-semibold text-slate-700">{entry.label || entry.sheet || 'Sheet export'}</div>
                                    <div className="text-[11px] text-slate-500">Sheet: {entry.sheet || '—'}</div>
                                    {entry.fileSuffix ? (
                                      <div className="text-[11px] text-slate-400">File suffix: {entry.fileSuffix}</div>
                                    ) : null}
                                  </div>
                                  <button
                                    type="button"
                                    disabled={busyDefinitionKey === definition.key}
                                    onClick={() => handleRemoveSheetExport(definition, idx)}
                                    className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-[11px] font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-60"
                                  >
                                    Remove
                                  </button>
                                </div>
                              ))
                            ) : (
                              <div className="rounded border border-slate-200 bg-white px-3 py-2 text-xs text-slate-500">No sheet exports yet. Sync from the workbook or add them manually.</div>
                            )
                          ) : (
                            <div className="rounded border border-slate-200 bg-white px-3 py-2 text-xs text-slate-500">Sheet exports are configured on the workbook template.</div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </section>
      </div>
    </div>
  );
}

export default TemplatesManager;

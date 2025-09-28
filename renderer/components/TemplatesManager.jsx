import React, { useCallback, useEffect, useMemo, useState } from 'react';
import ToastOverlay from './ToastOverlay';

const CATEGORY_LABELS = {
  client: 'Client details',
  event: 'Event details',
  venue: 'Venue details',
  financial: 'Financials',
  services: 'Services',
  other: 'Other'
};

const WORKBOOK_FILE_FILTERS = [{ name: 'Excel workbook', extensions: ['xlsx'] }];

function startCase(value) {
  return (value || '')
    .replace(/_/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\w/g, letter => letter.toUpperCase());
}

function slugify(value, fallback = 'workbook') {
  const base = (value || '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_{2,}/g, '_');
  return base || fallback;
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
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogMode, setDialogMode] = useState('create');
  const [dialogTarget, setDialogTarget] = useState(null);
  const [dialogLabel, setDialogLabel] = useState('');
  const [dialogFileSuffix, setDialogFileSuffix] = useState('');
  const [dialogRequiresTotal, setDialogRequiresTotal] = useState(true);
  const [dialogPath, setDialogPath] = useState('');
  const [dialogError, setDialogError] = useState('');
  const [dialogBusy, setDialogBusy] = useState(false);
  const [definitionSearch, setDefinitionSearch] = useState('');
  const [normalizeAllBusy, setNormalizeAllBusy] = useState(false);
  const [placeholdersCollapsed, setPlaceholdersCollapsed] = useState(false);
  const [templatesCollapsed, setTemplatesCollapsed] = useState(false);

  const restoreScroll = useCallback((position) => {
    if (typeof window === 'undefined') return;
    const top = Number.isFinite(position) ? position : 0;
    window.requestAnimationFrame(() => {
      window.scrollTo(0, top);
    });
  }, []);

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
      setDefinitions(allDefinitions);
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

    const merged = {
      ...definition,
      ...overrides,
      template_path: overrides.template_path !== undefined ? overrides.template_path : definition.template_path
    };

    const payload = {
      key: merged.key,
      doc_type: merged.doc_type,
      label: merged.label,
      description: merged.description,
      file_suffix: merged.file_suffix,
      invoice_variant: merged.invoice_variant,
      template_path: merged.template_path,
      requires_total: merged.requires_total ? 1 : 0,
      is_primary: merged.is_primary ? 1 : 0,
      is_active: merged.is_active === 0 ? 0 : 1,
      is_locked: merged.is_locked ? 1 : 0,
      sort_order: merged.sort_order
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

    try {
      const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
      const selectedPath = await api.chooseFile({
        title: `Select workbook for ${definition.label || startCase(definition.key)}`,
        filters: WORKBOOK_FILE_FILTERS
      });
      if (!selectedPath) return;

      setBusyDefinitionKey(definition.key);
      setMessage('Updating workbook template…');

      try {
        await api.normalizeTemplate?.({ templatePath: selectedPath });
      } catch (normalizeErr) {
        console.warn('Normalize template failed', normalizeErr);
      }

      await persistDefinition(definition, { template_path: selectedPath });
      setMessage(`${definition.label || startCase(definition.key)} template updated and normalized.`);
      await loadDefinitions();
      onTemplatesUpdated?.();
      restoreScroll(previousScroll);
    } catch (err) {
      console.error('Failed to replace template', err);
      setError(err?.message || 'Unable to update template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleClearTemplate = useCallback(async (definition) => {
    if (!definition) return;
    const confirmed = window.confirm('Clear the workbook template path? Documents will not generate until you pick a new file.');
    if (!confirmed) return;
    try {
      const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
      setBusyDefinitionKey(definition.key);
      await persistDefinition(definition, { template_path: '' });
      setMessage('Template cleared');
      await loadDefinitions();
      onTemplatesUpdated?.();
      restoreScroll(previousScroll);
    } catch (err) {
      console.error('Failed to clear template', err);
      setError(err?.message || 'Unable to clear template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleOpenCreateDialog = useCallback(() => {
    if (dialogBusy) return;
    setDialogMode('create');
    setDialogTarget(null);
    setDialogLabel('');
    
    setDialogFileSuffix('');
    setDialogRequiresTotal(true);
    setDialogPath('');
    setDialogError('');
    setDialogOpen(true);
  }, [dialogBusy]);

  const handleOpenEditDialog = useCallback((definition) => {
    if (!definition || dialogBusy) return;
    setDialogMode('edit');
    setDialogTarget(definition);
    setDialogLabel(definition.label || '');
    setDialogFileSuffix(definition.file_suffix || '');
    setDialogRequiresTotal(definition.requires_total ? 1 : 0);
    setDialogPath(definition.template_path || '');
    setDialogError('');
    setDialogOpen(true);
  }, [dialogBusy]);

  const handleCloseDialog = useCallback(() => {
    if (dialogBusy) return;
    setDialogOpen(false);
    setDialogMode('create');
    setDialogTarget(null);
    setDialogLabel('');
    setDialogFileSuffix('');
    setDialogRequiresTotal(true);
    setDialogPath('');
    setDialogError('');
  }, [dialogBusy]);

  const handleChooseDialogTemplate = useCallback(async () => {
    const api = window.api;
    if (!api || typeof api.chooseFile !== 'function') {
      setDialogError('File chooser unavailable');
      setTimeout(() => setDialogError(''), 2000);
      return;
    }
    try {
      const chosen = await api.chooseFile({
        title: 'Select workbook template',
        filters: WORKBOOK_FILE_FILTERS
      });
      if (chosen) {
        setDialogPath(chosen);
        setDialogError('');
      }
    } catch (err) {
      console.warn('Template selection cancelled or failed', err);
    }
  }, []);

  const handleSubmitDialog = useCallback(async (event) => {
    event.preventDefault();
    if (dialogBusy) return;
    const api = window.api;
    if (!api || typeof api.saveDocumentDefinition !== 'function') {
      setDialogError('Template API unavailable');
      return;
    }

    const label = dialogLabel.trim();
    if (!label) {
      setDialogError('Enter a template name');
      return;
    }

    let key;
    if (dialogMode === 'edit' && dialogTarget?.key) {
      key = dialogTarget.key;
    } else {
      const baseKey = slugify(label);
      key = baseKey;
      let suffix = 2;
      while (definitions.some(def => def.key === key)) {
        key = `${baseKey}_${suffix}`;
        suffix += 1;
      }
    }

    const docType = dialogTarget?.doc_type || 'workbook';

    const isEdit = dialogMode === 'edit' && dialogTarget;

    const payload = {
      key,
      doc_type: docType,
      label,
      description: '',
      file_suffix: dialogFileSuffix || '',
      invoice_variant: isEdit ? dialogTarget.invoice_variant : null,
      template_path: dialogPath || '',
      requires_total: dialogRequiresTotal ? 1 : 0,
      is_primary: isEdit ? (dialogTarget.is_primary ? 1 : 0) : definitions.length === 0 ? 1 : 0,
      is_active: isEdit ? (dialogTarget.is_active ? 1 : 0) : 1,
      is_locked: isEdit ? (dialogTarget.is_locked ? 1 : 0) : 0,
      sort_order: isEdit && Number.isFinite(dialogTarget.sort_order) ? dialogTarget.sort_order : definitions.length
    };

    try {
      setDialogBusy(true);
      setDialogError('');
      setMessage(dialogMode === 'edit' ? 'Template updated…' : 'Adding template…');
      await api.saveDocumentDefinition(business.id, payload);
      setMessage(dialogMode === 'edit' ? 'Template updated' : 'Template added');
      await loadDefinitions();
      onTemplatesUpdated?.();
      setDialogOpen(false);
      setDialogMode('create');
      setDialogTarget(null);
      setDialogLabel('');
      setDialogDocType('workbook');
      setDialogFileSuffix('');
      setDialogRequiresTotal(true);
      setDialogPath('');
    } catch (err) {
      console.error('Failed to save template', err);
      setDialogError(err?.message || 'Unable to save template');
    } finally {
      setDialogBusy(false);
      setTimeout(() => setMessage(''), 2500);
    }
  }, [dialogBusy, dialogLabel, dialogFileSuffix, dialogRequiresTotal, dialogPath, dialogMode, dialogTarget, business.id, definitions, loadDefinitions, onTemplatesUpdated]);

  const filteredDefinitions = useMemo(() => {
    const search = definitionSearch.trim().toLowerCase();
    if (!search) return definitions;
    return definitions.filter(definition => {
      const haystack = [
        definition.label,
        definition.key,
        definition.doc_type,
        definition.file_suffix,
        definition.description
      ].filter(Boolean).join(' ').toLowerCase();
      return haystack.includes(search);
    });
  }, [definitions, definitionSearch]);

  const handleDeleteDefinition = useCallback(async (definition) => {
    if (!definition) return;

    const current = definitions.find(def => def.key === definition.key) || definition;
    if (current.is_locked) {
      setError('This template is locked and cannot be deleted.');
      setTimeout(() => setError(''), 3000);
      return;
    }
    const confirmed = window.confirm(`Delete ${definition.label || definition.key}? This cannot be undone.`);
    if (!confirmed) return;

    const api = window.api;
    if (!api || typeof api.deleteDocumentDefinition !== 'function') {
      setError('Template API unavailable');
      setTimeout(() => setError(''), 2500);
      return;
    }

    try {
      const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
      setBusyDefinitionKey(definition.key);
      await api.deleteDocumentDefinition(business.id, definition.key);
      setMessage('Template deleted');
      await loadDefinitions();
      onTemplatesUpdated?.();
      restoreScroll(previousScroll);
    } catch (err) {
      console.error('Failed to delete template', err);
      setError(err?.message || 'Unable to delete template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [business.id, loadDefinitions, onTemplatesUpdated]);

  const handleToggleActive = useCallback(async (definition) => {
    if (!definition) return;
    const api = window.api;
    if (!api || typeof api.saveDocumentDefinition !== 'function') {
      setError('Template API unavailable');
      setTimeout(() => setError(''), 2500);
      return;
    }
    const nextActive = definition.is_active === 0 ? 1 : 0;
    try {
      const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
      setBusyDefinitionKey(definition.key);
      await persistDefinition({ ...definition, is_active: nextActive }, { template_path: definition.template_path });
      setMessage(nextActive ? 'Template activated' : 'Template deactivated');
      await loadDefinitions();
      onTemplatesUpdated?.();
      restoreScroll(previousScroll);
    } catch (err) {
      console.error('Failed to update template status', err);
      setError(err?.message || 'Unable to update template status');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

  const handleToggleLock = useCallback(async (definition) => {
    if (!definition) return;
    const api = window.api;
    if (!api || typeof api.saveDocumentDefinition !== 'function') {
      setError('Template API unavailable');
      setTimeout(() => setError(''), 2500);
      return;
    }
    const nextLocked = definition.is_locked ? 0 : 1;
    try {
      const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
      setBusyDefinitionKey(definition.key);
      await persistDefinition({ ...definition, is_locked: nextLocked }, { template_path: definition.template_path });
      setMessage(nextLocked ? 'Template locked' : 'Template unlocked');
      await loadDefinitions();
      onTemplatesUpdated?.();
      restoreScroll(previousScroll);
    } catch (err) {
      console.error('Failed to update template lock state', err);
      setError(err?.message || 'Unable to update template');
    } finally {
      setBusyDefinitionKey('');
      setTimeout(() => setMessage(''), 2500);
    }
  }, [loadDefinitions, onTemplatesUpdated, persistDefinition]);

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
      <ToastOverlay
        notices={[
          error ? { id: 'templates-error', tone: 'error', text: error } : null,
          message ? { id: 'templates-message', tone: 'info', text: message } : null,
          copyFeedback ? { id: 'templates-copy', tone: 'success', text: copyFeedback } : null
        ]}
      />

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
            <header className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
              <div>
                <h2 className="text-sm font-semibold text-slate-700">Placeholders</h2>
                <p className="mt-1 text-xs text-slate-500">Copy tokens and see where their data comes from.</p>
              </div>
              <button
                type="button"
                onClick={() => setPlaceholdersCollapsed(prev => !prev)}
                className="inline-flex items-center gap-1 rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
                aria-expanded={!placeholdersCollapsed}
              >
                {placeholdersCollapsed ? 'Expand' : 'Collapse'}
              </button>
            </header>
            {!placeholdersCollapsed ? (
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
            ) : null}
          </div>

          <div className="rounded border border-slate-200 bg-white shadow-sm">
            <header className="border-b border-slate-200 px-4 py-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h2 className="text-sm font-semibold text-slate-700">Templates</h2>
                <p className="mt-1 text-xs text-slate-500">Manage all document templates and quickly access related actions.</p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => setTemplatesCollapsed(prev => !prev)}
                  className="inline-flex items-center gap-1 rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
                  aria-expanded={!templatesCollapsed}
                >
                  {templatesCollapsed ? 'Expand' : 'Collapse'}
                </button>
                <button
                  type="button"
                  onClick={handleOpenCreateDialog}
                  disabled={dialogBusy}
                  className="inline-flex items-center rounded border border-indigo-200 bg-indigo-50 px-3 py-1.5 text-xs font-medium text-indigo-700 hover:bg-indigo-100 disabled:opacity-60"
                >
                  {dialogBusy && dialogMode === 'create' ? 'Adding…' : 'Add template'}
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    if (normalizeAllBusy) return;
                    const api = window.api;
                    if (!api || typeof api.normalizeTemplate !== 'function') {
                      setError('Normalize API unavailable');
                      setTimeout(() => setError(''), 2500);
                      return;
                    }
                    const paths = definitions
                      .map(def => def.template_path)
                      .filter(path => typeof path === 'string' && path.trim());
                    if (!paths.length) {
                      setMessage('No templates to normalize');
                      setTimeout(() => setMessage(''), 2000);
                      return;
                    }
                    const previousScroll = typeof window !== 'undefined' ? window.scrollY : 0;
                    setNormalizeAllBusy(true);
                    setMessage('Normalizing templates…');
                    try {
                      await Promise.all(paths.map(path => api.normalizeTemplate({ templatePath: path })));            
                      setMessage('Templates normalized');
                    } catch (err) {
                      console.error('Normalize all failed', err);
                      setError(err?.message || 'Unable to normalize templates');
                    } finally {
                      setNormalizeAllBusy(false);
                      setTimeout(() => setMessage(''), 2500);
                      restoreScroll(previousScroll);
                    }
                  }}
                  disabled={normalizeAllBusy || !definitions.some(def => def.template_path)}
                  className="inline-flex items-center rounded border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                >
                  {normalizeAllBusy ? 'Normalizing…' : 'Normalize all'}
                </button>
              </div>
            </header>
            {!templatesCollapsed ? (
              <div className="px-4 py-3 space-y-3">
                {loadingDefinitions ? (
                  <div className="text-sm text-slate-500">Loading templates…</div>
                ) : !definitions.length ? (
                  <div className="text-sm text-slate-500">No document types configured yet.</div>
                ) : (
                  <>
                    <div className="grid gap-2">
                      <label className="flex items-center gap-2 text-sm text-slate-600">
                        <span className="sr-only">Search templates</span>
                        <input
                          type="search"
                          value={definitionSearch}
                          onChange={event => setDefinitionSearch(event.target.value)}
                          placeholder="Search templates"
                          className="w-full rounded border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                        />
                      </label>
                    </div>
                    {!filteredDefinitions.length ? (
                      <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-500">
                        No templates match your filters.
                      </div>
                    ) : null}
                    <div className="space-y-2">
                      {filteredDefinitions.map(definition => {
                      const templatePath = definition.template_path || '';
                      const isLocked = definition.is_locked === 1;
                      const isInactive = definition.is_active === 0;

                      const actionButtons = [
                        {
                          label: 'Edit template',
                          icon: '✏️',
                          onClick: () => handleOpenEditDialog(definition),
                          disabled: dialogBusy
                        },
                        {
                          label: 'Replace file',
                          icon: '📂',
                          onClick: () => handleReplaceTemplate(definition),
                          disabled: busyDefinitionKey === definition.key
                        },
                        {
                          label: 'Open file',
                          icon: '🔍',
                          onClick: () => handleOpenTemplate(definition),
                          disabled: !templatePath
                        },
                        {
                          label: 'Clear path',
                          icon: '🧹',
                          onClick: () => handleClearTemplate(definition),
                          disabled: busyDefinitionKey === definition.key || !templatePath
                        },
                        {
                          label: definition.is_active === 0 ? 'Activate' : 'Deactivate',
                          icon: definition.is_active === 0 ? '✅' : '🚫',
                          onClick: () => handleToggleActive(definition),
                          disabled: busyDefinitionKey === definition.key
                        },
                        {
                          label: isLocked ? 'Unlock' : 'Lock',
                          icon: isLocked ? '🔓' : '🔒',
                          onClick: () => handleToggleLock(definition),
                          disabled: busyDefinitionKey === definition.key
                        },
                        {
                          label: 'Delete template',
                          icon: '🗑️',
                          onClick: () => handleDeleteDefinition(definition),
                          disabled: busyDefinitionKey === definition.key
                        }
                      ];

                      return (
                        <div
                          key={definition.key}
                          className="rounded border border-slate-200 bg-white px-3 py-2 shadow-sm transition hover:border-indigo-200 hover:shadow"
                        >
                          <div className="flex flex-wrap items-center justify-between gap-3">
                            <div className="min-w-0 space-y-1">
                              <div className="flex items-center gap-2">
                                <span className="text-sm font-semibold text-slate-800 truncate" title={definition.label || definition.key}>
                                  {definition.label || startCase(definition.key)}
                                </span>
                                
                                {isLocked ? <span className="text-[11px] text-amber-600" title="Locked">🔒</span> : null}
                                {isInactive ? <span className="text-[11px] text-rose-600" title="Inactive">⏸</span> : null}
                              </div>
                              <div className="text-xs text-slate-500 truncate" title={templatePath || 'No template selected'}>
                                {templatePath || 'No template selected yet'}
                              </div>
                              <div className="flex flex-wrap gap-3 text-[11px] text-slate-500">
                                <span>{definition.requires_total ? 'Requires totals' : 'No totals required'}</span>
                                {definition.file_suffix ? <span title="File suffix">Suffix: {definition.file_suffix}</span> : null}
                              </div>
                            </div>
                            <div className="flex flex-wrap items-center gap-2">
                              {actionButtons.map(action => (
                                <button
                                  key={action.label}
                                  type="button"
                                  onClick={action.onClick}
                                  disabled={action.disabled}
                                  className="inline-flex h-9 w-9 items-center justify-center rounded border border-slate-300 text-lg hover:bg-slate-100 disabled:opacity-60"
                                  title={action.label}
                                  aria-label={action.label}
                                >
                                  <span role="img" aria-hidden="true">{action.icon}</span>
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                      );
                      })}
                    </div>
                  </>
                )}
              </div>
            ) : null}
          </div>
        </section>
      </div>
      {dialogOpen ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 px-4 py-6">
          <div className="w-full max-w-md rounded-lg bg-white shadow-xl">
            <form onSubmit={handleSubmitDialog} className="space-y-4 p-5">
              <div className="flex items-start justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-slate-800">{dialogMode === 'edit' ? 'Edit template' : 'Add template'}</h3>
                  <p className="text-sm text-slate-500">Set the metadata and optional file path for this template.</p>
                </div>
                <button
                  type="button"
                  onClick={handleCloseDialog}
                  className="text-slate-400 transition hover:text-slate-600"
                  aria-label="Close"
                  disabled={dialogBusy}
                >
                  ✕
                </button>
              </div>

              {dialogError ? (
                <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-600">{dialogError}</div>
              ) : null}

              <label className="block text-sm font-medium text-slate-700">
                Template name
                <input
                  type="text"
                  value={dialogLabel}
                  onChange={event => setDialogLabel(event.target.value)}
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                  placeholder="e.g. Client Pack"
                  disabled={dialogBusy}
                  required
                />
              </label>

              <label className="block text-sm font-medium text-slate-700">
                File suffix (optional)
                <input
                  type="text"
                  value={dialogFileSuffix}
                  onChange={event => setDialogFileSuffix(event.target.value)}
                  className="mt-1 w-full rounded border border-slate-300 px-3 py-2 text-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                  placeholder="e.g. - Quote"
                  disabled={dialogBusy}
                />
              </label>

              <label className="flex items-center gap-2 text-sm text-slate-600">
                <input
                  type="checkbox"
                  checked={Boolean(dialogRequiresTotal)}
                  onChange={event => setDialogRequiresTotal(event.target.checked)}
                  disabled={dialogBusy}
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                />
                Requires totals
              </label>

              <div className="space-y-1">
                <div className="flex items-center justify-between text-sm font-medium text-slate-700">
                  <span>Template file (optional)</span>
                  <button
                    type="button"
                    onClick={handleChooseDialogTemplate}
                    disabled={dialogBusy}
                    className="inline-flex items-center rounded border border-slate-300 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                  >
                    Choose file
                  </button>
                </div>
                <div className="rounded border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-500 break-all min-h-[2.5rem]">
                  {dialogPath || 'No file selected. You can choose this later.'}
                </div>
              </div>

              <div className="flex justify-end gap-2">
                <button
                  type="button"
                  onClick={handleCloseDialog}
                  disabled={dialogBusy}
                  className="inline-flex items-center rounded border border-slate-300 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-60"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={dialogBusy}
                  className="inline-flex items-center rounded bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-500 disabled:opacity-60"
                >
                  {dialogBusy ? 'Saving…' : 'Save template'}
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
    </div>
  );
}

export default TemplatesManager;

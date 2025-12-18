import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import ToastOverlay from './ToastOverlay';

const CATEGORY_LABELS = {
  client: 'Client details',
  event: 'Event details',
  venue: 'Venue details',
  financial: 'Financials',
  services: 'Services',
  other: 'Other'
};

const WORKBOOK_FILE_FILTERS = [{ name: 'Excel workbook or PDF', extensions: ['xlsx', 'pdf'] }];

const TEMPLATES_MANAGER_TABS = [
  { key: 'templates', label: 'Templates' },
  { key: 'placeholders', label: 'Placeholders' }
];

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
  const MISSING_TPL_NOTICE_KEY = 'invoiceMaster:templatesMissingNotice:v1';
  const [placeholders, setPlaceholders] = useState([]);
  const [valueSources, setValueSources] = useState({});
  const [definitions, setDefinitions] = useState([]);
  const [reorderBusy, setReorderBusy] = useState(false);
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
  const [dialogPath, setDialogPath] = useState('');
  const [dialogError, setDialogError] = useState('');
  const [dialogBusy, setDialogBusy] = useState(false);
  const [definitionSearch, setDefinitionSearch] = useState('');
  const [debouncedDefinitionSearch, setDebouncedDefinitionSearch] = useState('');
  const [normalizeAllBusy, setNormalizeAllBusy] = useState(false);
  const [placeholdersCollapsed, setPlaceholdersCollapsed] = useState(false);
  const [templatesCollapsed, setTemplatesCollapsed] = useState(false);
  const [activeTab, setActiveTab] = useState('templates');
  const [menuOpenKey, setMenuOpenKey] = useState('');
  const menuRef = useRef(null);

  // Persist selected tab between sessions
  useEffect(() => {
    try {
      const stored = window.localStorage.getItem('templatesManager:activeTab');
      if (stored && (stored === 'templates' || stored === 'placeholders')) {
        setActiveTab(stored);
      }
    } catch (_err) {}
  }, []);

  useEffect(() => {
    try {
      window.localStorage.setItem('templatesManager:activeTab', activeTab);
    } catch (_err) {}
  }, [activeTab]);

  // Close kebab menu on outside click or Escape
  useEffect(() => {
    const handleDocClick = (event) => {
      if (!menuOpenKey) return;
      const el = menuRef.current;
      if (el && el.contains(event.target)) return;
      setMenuOpenKey('');
    };
    const handleKeyDown = (event) => {
      if (event.key === 'Escape' && menuOpenKey) setMenuOpenKey('');
    };
    document.addEventListener('mousedown', handleDocClick);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handleDocClick);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [menuOpenKey]);

  useEffect(() => {
    if (!menuOpenKey) return;
    // focus first menu item when menu opens
    const id = requestAnimationFrame(() => {
      const first = menuRef.current?.querySelector('button[role="menuitem"]');
      if (first && typeof first.focus === 'function') first.focus();
    });
    return () => cancelAnimationFrame(id);
  }, [menuOpenKey]);

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

      // One-time notice if any definitions are missing a template path
      try {
        const seen = window.localStorage.getItem(MISSING_TPL_NOTICE_KEY);
        if (seen !== '1') {
          const missingCount = allDefinitions.filter(d => !d?.template_path).length;
          if (missingCount > 0) {
            setMessage(`${missingCount} template${missingCount === 1 ? '' : 's'} need a workbook path. Use “Replace” to select a file.`);
            window.localStorage.setItem(MISSING_TPL_NOTICE_KEY, '1');
          }
        }
      } catch (_err) {}
    } catch (err) {
      console.error('Failed to load document definitions', err);
      setError(err?.message || 'Unable to load templates');
    } finally {
      setLoadingDefinitions(false);
    }
  }, [business.id]);

  useEffect(() => {
    const id = setTimeout(() => setDebouncedDefinitionSearch(definitionSearch.trim().toLowerCase()), 200);
    return () => clearTimeout(id);
  }, [definitionSearch]);

  useEffect(() => {
    // restore UI state
    try {
      const ph = window.localStorage.getItem('templatesManager:placeholdersCollapsed');
      if (ph != null) setPlaceholdersCollapsed(ph === '1');
      const tm = window.localStorage.getItem('templatesManager:templatesCollapsed');
      if (tm != null) setTemplatesCollapsed(tm === '1');
      const lastCat = window.localStorage.getItem('templatesManager:lastPlaceholderCategory');
      if (lastCat) setActiveCategory(lastCat);
    } catch (_err) {}
    loadPlaceholders();
  }, [loadPlaceholders]);

  useEffect(() => {
    loadDefinitions();
  }, [loadDefinitions]);

  useEffect(() => {
    try {
      window.localStorage.setItem('templatesManager:placeholdersCollapsed', placeholdersCollapsed ? '1' : '0');
    } catch (_err) {}
  }, [placeholdersCollapsed]);

  useEffect(() => {
    try {
      window.localStorage.setItem('templatesManager:templatesCollapsed', templatesCollapsed ? '1' : '0');
    } catch (_err) {}
  }, [templatesCollapsed]);

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
      invoice_variant: merged.invoice_variant,
      template_path: merged.template_path,
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

  useEffect(() => {
    if (!activeCategory) return;
    try { window.localStorage.setItem('templatesManager:lastPlaceholderCategory', activeCategory); } catch (_err) {}
  }, [activeCategory]);

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
        title: `Select template for ${definition.label || startCase(definition.key)}`,
        filters: WORKBOOK_FILE_FILTERS
      });
      if (!selectedPath) return;

      setBusyDefinitionKey(definition.key);
      setMessage('Updating template…');

      try {
        const lowerPath = String(selectedPath).toLowerCase();
        if (lowerPath.endsWith('.xlsx')) {
          await api.normalizeTemplate?.({ templatePath: selectedPath });
        }
      } catch (normalizeErr) {
        console.warn('Normalize template failed', normalizeErr);
      }

      await persistDefinition(definition, { template_path: selectedPath });
      setMessage(`${definition.label || startCase(definition.key)} template updated${String(selectedPath).toLowerCase().endsWith('.xlsx') ? ' and normalized' : ''}.`);
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
    
    // totals no longer required; no state to manage
    setDialogPath('');
    setDialogError('');
    setDialogOpen(true);
  }, [dialogBusy]);

  const handleOpenEditDialog = useCallback((definition) => {
    if (!definition || dialogBusy) return;
    setDialogMode('edit');
    setDialogTarget(definition);
    setDialogLabel(definition.label || '');
    // totals are irrelevant for editing now
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
    // reset not needed for totals
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
      invoice_variant: isEdit ? dialogTarget.invoice_variant : null,
      template_path: dialogPath || '',
      // requires_total removed
      is_primary: 0,
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
      setDialogPath('');
    } catch (err) {
      console.error('Failed to save template', err);
      setDialogError(err?.message || 'Unable to save template');
    } finally {
      setDialogBusy(false);
      setTimeout(() => setMessage(''), 2500);
    }
  }, [dialogBusy, dialogLabel, dialogPath, dialogMode, dialogTarget, business.id, definitions, loadDefinitions, onTemplatesUpdated]);

  const filteredDefinitions = useMemo(() => {
    const search = debouncedDefinitionSearch;
    if (!search) return definitions;
    return definitions.filter(definition => {
      const haystack = [
        definition.label,
        definition.key,
        definition.doc_type,
        definition.description
      ].filter(Boolean).join(' ').toLowerCase();
      return haystack.includes(search);
    });
  }, [definitions, debouncedDefinitionSearch]);

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

  const moveDefinition = useCallback(async (key, delta) => {
    if (!Array.isArray(definitions) || !definitions.length) return;
    const index = definitions.findIndex(d => d.key === key);
    if (index === -1) return;
    const nextIndex = index + delta;
    if (nextIndex < 0 || nextIndex >= definitions.length) return;
    const updated = [...definitions];
    const [moved] = updated.splice(index, 1);
    updated.splice(nextIndex, 0, moved);
    setDefinitions(updated);
    try {
      setReorderBusy(true);
      const keys = updated.map(d => d.key);
      await window.api?.reorderDocumentDefinitions?.(business.id, keys);
      setMessage('Order updated');
      setTimeout(() => setMessage(''), 1500);
    } catch (err) {
      console.error('Failed to reorder templates', err);
      setError(err?.message || 'Unable to reorder templates');
    } finally {
      setReorderBusy(false);
    }
  }, [definitions, business.id]);

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

      <div className="border-b border-slate-200">
        <nav className="flex gap-2" aria-label="Templates manager sections" role="tablist">
          {TEMPLATES_MANAGER_TABS.map(tab => {
            const isActive = activeTab === tab.key;
            const tabId = `templates-tab-${tab.key}`;
            const panelId = `templates-panel-${tab.key}`;
            return (
              <button
                key={tab.key}
                id={tabId}
                type="button"
                onClick={() => setActiveTab(tab.key)}
                className={`inline-flex items-center rounded-t-md border px-4 py-2 text-sm font-medium transition focus:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500 ${isActive ? 'border-slate-200 border-b-white bg-white text-indigo-700 shadow-sm' : 'border-transparent text-slate-500 hover:border-slate-200 hover:text-slate-700'}`}
                aria-selected={isActive}
                aria-controls={panelId}
                role="tab"
              >
                {tab.label}
              </button>
            );
          })}
        </nav>
      </div>

      {activeTab === 'placeholders' ? (
        <section
          id="templates-panel-placeholders"
          role="tabpanel"
          aria-labelledby="templates-tab-placeholders"
          className="space-y-6"
        >
          <div className="flex flex-col gap-6 lg:flex-row">
            <aside className="lg:w-64 flex-shrink-0 space-y-2">
              {loading ? (
                <div className="rounded border border-slate-200 bg-white px-3 py-2 text-sm text-slate-500">Loading placeholders…</div>
              ) : categoryEntries.length === 0 ? (
                <div className="rounded border border-slate-200 bg-white px-3 py-2 text-sm text-slate-500">No placeholders found.</div>
              ) : (
                categoryEntries.map(category => {
                  const isActiveCategory = activeCategory === category.key;
                  return (
                    <button
                      key={category.key}
                      type="button"
                      onClick={() => setActiveCategory(category.key)}
                      className={`flex w-full items-center justify-between rounded-lg border px-3 py-2 text-sm transition ${isActiveCategory ? 'border-indigo-200 bg-indigo-50 text-indigo-700 font-semibold' : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'}`}
                    >
                      <span>{category.label}</span>
                      <span className={`inline-flex h-5 min-w-[1.5rem] items-center justify-center rounded-full text-xs ${isActiveCategory ? 'bg-indigo-100 text-indigo-700' : 'bg-slate-100 text-slate-500'}`}>
                        {category.rows.length}
                      </span>
                    </button>
                  );
                })
              )}
            </aside>

            <div className="flex-1">
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
            </div>
          </div>
        </section>
      ) : null}

      {activeTab === 'templates' ? (
        <section
          id="templates-panel-templates"
          role="tabpanel"
          aria-labelledby="templates-tab-templates"
          className="rounded border border-slate-200 bg-white shadow-sm"
        >
          <header className="border-b border-slate-200 px-4 py-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h2 className="text-sm font-semibold text-slate-700">Templates</h2>
              <p className="mt-1 text-xs text-slate-500">Manage and reorder document templates.</p>
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
                    .filter(path => typeof path === 'string' && path.trim())
                    .filter(path => path.toLowerCase().endsWith('.xlsx'));
                  if (!paths.length) {
                    setMessage('No Excel templates to normalize');
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

                      // actions moved into kebab menu

                      const canReorder = !debouncedDefinitionSearch && !reorderBusy;
                      const currentIndex = definitions.findIndex(d => d.key === definition.key);
                      const isFirst = currentIndex <= 0;
                      const isLast = currentIndex === definitions.length - 1;
                      return (
                        <div
                          key={definition.key}
                          className={`relative rounded border bg-white px-3 py-2 shadow-sm transition hover:border-indigo-200 hover:shadow border-slate-200`}
                        >
                          <div className="flex items-start justify-between gap-3">
                                <div className="flex items-start gap-2 min-w-0">
                                  {canReorder ? (
                                <div className="flex flex-col items-center gap-1 w-7 flex-none" aria-hidden="false">
                                  <button
                                    type="button"
                                    onClick={() => moveDefinition(definition.key, -1)}
                                    disabled={reorderBusy || isFirst}
                                    className="inline-flex h-5 w-5 items-center justify-center rounded border border-slate-100 bg-white text-[11px] text-slate-200 opacity-60 hover:opacity-100 hover:text-slate-700 hover:border-slate-300 hover:bg-slate-100 disabled:opacity-40"
                                    title="Move up"
                                    aria-label="Move up"
                                  >
                                    ▲
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => moveDefinition(definition.key, 1)}
                                    disabled={reorderBusy || isLast}
                                    className="inline-flex h-5 w-5 items-center justify-center rounded border border-slate-100 bg-white text-[11px] text-slate-200 opacity-60 hover:opacity-100 hover:text-slate-700 hover:border-slate-300 hover:bg-slate-100 disabled:opacity-40"
                                    title="Move down"
                                    aria-label="Move down"
                                  >
                                    ▼
                                  </button>
                                </div>
                              ) : null}
                              <div className="min-w-0">
                                <div className="flex items-center gap-2">
                                  <span className="text-sm font-semibold text-slate-800 truncate" title={definition.label || definition.key}>
                                    {definition.label || startCase(definition.key)}
                                  </span>
                                  {isInactive ? (
                                    <span className="inline-flex items-center rounded-full border border-amber-200 bg-amber-50 px-2 py-0.5 text-[11px] text-amber-700" title="Inactive">
                                      Inactive
                                    </span>
                                  ) : null}
                                  {isLocked ? (
                                    <span className="text-[11px] text-amber-600 align-middle" title="Locked">🔒</span>
                                  ) : null}
                                </div>
                                <div className="text-xs text-slate-500 truncate mt-0.5" title={templatePath || 'No template selected'}>
                                  {!templatePath ? (
                                    <span className="inline-flex items-center gap-1 rounded-full border border-rose-200 bg-rose-50 px-2 py-0.5 text-[11px] text-rose-700 mr-2">Missing file</span>
                                  ) : null}
                                  {templatePath || 'No template selected yet'}
                                </div>
                              </div>
                            </div>
                            <div className="relative ml-auto">
                              <button
                                type="button"
                                onClick={() => setMenuOpenKey(menuOpenKey === definition.key ? '' : definition.key)}
                                className="inline-flex h-9 w-9 items-center justify-center rounded border border-slate-300 text-lg hover:bg-slate-100"
                                aria-haspopup="menu"
                                aria-expanded={menuOpenKey === definition.key}
                                aria-label="More actions"
                              >
                                ⋯
                              </button>
                              {menuOpenKey === definition.key ? (
                                <div
                                  ref={menuRef}
                                  className="absolute right-0 z-20 mt-2 w-52 rounded border border-slate-200 bg-white p-1 shadow-lg"
                                  role="menu"
                                  onKeyDown={(e) => {
                                    const items = Array.from(menuRef.current?.querySelectorAll('button[role="menuitem"]') || []);
                                    if (!items.length) return;
                                    const idx = items.indexOf(document.activeElement);
                                    if (e.key === 'ArrowDown') {
                                      e.preventDefault();
                                      const next = items[(idx + 1 + items.length) % items.length];
                                      next?.focus();
                                    } else if (e.key === 'ArrowUp') {
                                      e.preventDefault();
                                      const prev = items[(idx - 1 + items.length) % items.length];
                                      prev?.focus();
                                    } else if (e.key === 'Home') {
                                      e.preventDefault(); items[0]?.focus();
                                    } else if (e.key === 'End') {
                                      e.preventDefault(); items[items.length - 1]?.focus();
                                    }
                                  }}
                                >
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleOpenEditDialog(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100"
                                    disabled={dialogBusy}
                                    role="menuitem"
                                  >
                                    Edit template
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleReplaceTemplate(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={busyDefinitionKey === definition.key}
                                    role="menuitem"
                                  >
                                    Replace file
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleOpenTemplate(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={!templatePath}
                                    role="menuitem"
                                  >
                                    Open file
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); window.api?.showItemInFolder?.(templatePath); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={!templatePath}
                                    role="menuitem"
                                  >
                                    Reveal in Finder
                                  </button>
                                  <button
                                    type="button"
                                    onClick={async () => {
                                      try {
                                        if (templatePath && navigator?.clipboard?.writeText) {
                                          await navigator.clipboard.writeText(templatePath);
                                          setCopyFeedback('Template path copied');
                                          setTimeout(() => setCopyFeedback(''), 2000);
                                        }
                                      } catch (_err) {}
                                      setMenuOpenKey('');
                                    }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={!templatePath}
                                    role="menuitem"
                                  >
                                    Copy path
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleClearTemplate(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={busyDefinitionKey === definition.key || !templatePath}
                                    role="menuitem"
                                  >
                                    Clear path
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleToggleActive(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={busyDefinitionKey === definition.key}
                                    role="menuitem"
                                  >
                                    {definition.is_active === 0 ? 'Activate' : 'Deactivate'}
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleToggleLock(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-slate-100 disabled:opacity-60"
                                    disabled={busyDefinitionKey === definition.key}
                                    role="menuitem"
                                  >
                                    {isLocked ? 'Unlock' : 'Lock'}
                                  </button>
                                  <div className="my-1 border-t border-slate-200" />
                                  <button
                                    type="button"
                                    onClick={() => { setMenuOpenKey(''); handleDeleteDefinition(definition); }}
                                    className="block w-full rounded px-3 py-1.5 text-left text-sm text-rose-700 hover:bg-rose-50 disabled:opacity-60"
                                    disabled={busyDefinitionKey === definition.key}
                                    role="menuitem"
                                  >
                                    Delete template
                                  </button>
                                </div>
                              ) : null}
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
        </section>
      ) : null}

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

              {/* Suffix field removed */}

              {/* Requires totals removed */}

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

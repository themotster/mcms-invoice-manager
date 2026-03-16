import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';

function IconButton({ label, onClick, disabled, className = '', children, size = 'md' }) {
  const sizePx = size === 'sm' ? 28 : (size === 'lg' ? 40 : 36);
  return (
    <button
      type="button"
      onClick={(e) => { e.stopPropagation(); onClick?.(e); }}
      disabled={disabled}
      style={{
        width: sizePx, height: sizePx, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
        border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569', cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.6 : 1
      }}
      aria-label={label}
      title={label}
    >
      {children}
    </button>
  );
}
function EyeIcon({ style }) {
  return (
    <svg style={{ width: 16, height: 16, ...style }} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M2.25 12s2.75-6.75 9.75-6.75 9.75 6.75 9.75 6.75-2.75 6.75-9.75 6.75S2.25 12 2.25 12Z" />
      <path d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z" />
    </svg>
  );
}
function RevealIcon({ style }) {
  return (
    <svg style={{ width: 16, height: 16, ...style }} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 6.25A2.25 2.25 0 0 1 6.25 4h4.086c.414 0 .812.165 1.105.459L13.5 6.5H19A2 2 0 0 1 21 8.5V9H4V6.25Z" fill="currentColor" opacity="0.5" />
      <path d="M3 9.75A1.75 1.75 0 0 1 4.75 8h15.5A1.75 1.75 0 0 1 22 9.75v7.5A2.75 2.75 0 0 1 19.25 20H6A3 3 0 0 1 3 17V9.75Z" fill="currentColor" />
    </svg>
  );
}
function DeleteIcon({ style }) {
  return (
    <svg style={{ width: 16, height: 16, ...style }} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M6 7h12" /><path d="M9.5 7V5.75A1.75 1.75 0 0 1 11.25 4h1.5A1.75 1.75 0 0 1 14.5 5.75V7" /><path d="M17 7v10.25A1.75 1.75 0 0 1 15.25 19h-6.5A1.75 1.75 0 0 1 7 17.25V7" /><path d="M10 11v5" /><path d="M14 11v5" />
    </svg>
  );
}
function PencilIcon({ style }) {
  return (
    <svg style={{ width: 16, height: 16, ...style }} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M16.5 3.5L20.5 7.5L7 21H3V17L16.5 3.5Z" /><path d="M15 5L19 9" />
    </svg>
  );
}

function Spinner({ size = 20, style = {} }) {
  return (
    <span
      role="status"
      aria-label="Loading"
      style={{
        display: 'inline-block',
        width: size,
        height: size,
        border: '2px solid #e2e8f0',
        borderTopColor: '#4f46e5',
        borderRadius: '50%',
        animation: 'mcms-spin 0.7s linear infinite',
        ...style
      }}
    />
  );
}

function App() {
  const BUSINESS_ID = 1; // MCMS
  const [activeTab, setActiveTab] = useState('invoices'); // 'invoices' | 'contacts' | 'templates'
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');

  // Unified client box state
  const [clientQuery, setClientQuery] = useState('');
  const [clientFocus, setClientFocus] = useState(false);
  const [selectedClient, setSelectedClient] = useState(null); // full client object
  const [excelTemplatePath, setExcelTemplatePath] = useState('');
  const [excelBusy, setExcelBusy] = useState(false);
  // Delete modal state
  const [deleteModalOpen, setDeleteModalOpen] = useState(false);
  const [deleteModalData, setDeleteModalData] = useState({ selected: null });
  // Placeholders UI removed (Excel-first flow)

  const [docs, setDocs] = useState([]);
  const [docsLoading, setDocsLoading] = useState(true);
  // Sorting for invoice log
  const [sortKey, setSortKey] = useState('created'); // 'invoice' | 'client' | 'created'
  const [sortDir, setSortDir] = useState('desc'); // 'asc' | 'desc'

  const sortedDocs = useMemo(() => {
    const items = Array.isArray(docs) ? [...docs] : [];
    const getVal = (d) => {
      switch (sortKey) {
        case 'invoice': return Number(d?.number) || 0;
        case 'client': return String(d?.display_client_name || d?.client_name || '').toLowerCase();
        case 'created': default: {
          const v = String(d?.event_date || d?.document_date || d?.created_at || '');
          return v; // ISO-like strings compare lexicographically
        }
      }
    };
    items.sort((a, b) => {
      const va = getVal(a);
      const vb = getVal(b);
      if (va === vb) return 0;
      if (sortDir === 'asc') return va > vb ? 1 : -1;
      return va < vb ? 1 : -1;
    });
    return items;
  }, [docs, sortKey, sortDir]);

  const toggleSort = (key) => {
    setSortKey(prevKey => {
      if (prevKey === key) {
        setSortDir(prevDir => (prevDir === 'asc' ? 'desc' : 'asc'));
        return prevKey;
      }
      setSortDir('asc');
      return key;
    });
  };
  const [fieldValues, setFieldValues] = useState({}); // key -> value
  const [invoiceNumber, setInvoiceNumber] = useState('');
  const [invoiceNumTouched, setInvoiceNumTouched] = useState(false);
  const [invoiceNumTaken, setInvoiceNumTaken] = useState(false);
  const [invoiceNumChecking, setInvoiceNumChecking] = useState(false);
  const [invoiceNumError, setInvoiceNumError] = useState('');
  const [savePath, setSavePath] = useState('');
  const [dbPath, setDbPath] = useState('');

  // Create invoice form state
  const todayISO = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const [invoiceDate, setInvoiceDate] = useState(todayISO);
  const [dueDate, setDueDate] = useState('On receipt');
  const [lineItems, setLineItems] = useState([{ date: todayISO, description: '', amount: '' }]);
  const [totalOverride, setTotalOverride] = useState('');
  const [amountReceived, setAmountReceived] = useState('');
  const [discountDescription, setDiscountDescription] = useState('');
  const [discountAmount, setDiscountAmount] = useState('');
  const [createBusy, setCreateBusy] = useState(false);
  const [invoiceModalOpen, setInvoiceModalOpen] = useState(false);
  const [invoiceModalMode, setInvoiceModalMode] = useState('new'); // 'new' | 'edit'
  const [editingDocument, setEditingDocument] = useState(null); // when edit, the doc row

  // Contacts editor state
  const [contactQuery, setContactQuery] = useState('');
  const [contactFocus, setContactFocus] = useState(false);
  const [contactSelectedId, setContactSelectedId] = useState(null);
  const [contactDetails, setContactDetails] = useState({ client: null, emails: [], phones: [], addresses: [] });
  const [contactBusy, setContactBusy] = useState(false);
  const [contactModalOpen, setContactModalOpen] = useState(false);

  const refreshClients = useCallback(async () => {
    setLoading(true); setError('');
    try {
      const list = await window.api?.getClients?.();
      const filtered = Array.isArray(list) ? list.filter(c => !c.business_id || c.business_id === BUSINESS_ID) : [];
      setClients(filtered);
    } catch (err) {
      setError(err?.message || 'Unable to load clients');
    } finally { setLoading(false); }
  }, []);

  const loadClientDetails = useCallback(async (clientId) => {
    if (!clientId) { setContactDetails({ client: null, emails: [], phones: [], addresses: [] }); return; }
    try {
      const det = await window.api?.getClientDetails?.(clientId);
      if (!det || !det.client) { setContactDetails({ client: null, emails: [], phones: [], addresses: [] }); return; }
      setContactDetails({
        client: det.client,
        emails: Array.isArray(det.emails) ? det.emails : [],
        phones: Array.isArray(det.phones) ? det.phones : [],
        addresses: Array.isArray(det.addresses) ? det.addresses : []
      });
      setContactSelectedId(clientId);
    } catch (err) {
      setError(err?.message || 'Unable to load contact');
    }
  }, []);

  const refreshDocs = useCallback(async () => {
    setDocsLoading(true); setError('');
    try {
      // Reconcile DB with filesystem first to avoid phantom entries
      try { await window.api?.cleanOrphanDocuments?.({ businessId: BUSINESS_ID }); } catch (_) {}
      const items = await window.api?.getDocuments?.({ businessId: BUSINESS_ID, docType: 'invoice' });
      const list = Array.isArray(items) ? items : [];
      // Include all invoices from DB (includeMissing: true) so log isn't blank if paths changed or files moved
      let enriched = list;
      try {
        enriched = await window.api?.filterDocumentsByExistingFiles?.(list, { includeMissing: true });
      } catch (_) {}
      setDocs(Array.isArray(enriched) ? enriched : list);
    } catch (err) {
      setError(err?.message || 'Unable to load invoices');
      setDocs([]);
    } finally { setDocsLoading(false); }
  }, []);

  const loadInvoiceDefinition = useCallback(async () => {
    try {
      const defs = await window.api.getDocumentDefinitions(BUSINESS_ID, { includeInactive: true });
      const list = Array.isArray(defs) ? defs : [];
      const def = list.find(d => String(d.key || '').toLowerCase() === 'invoice_balance');
      setExcelTemplatePath(def?.template_path || '');
      if (def?.template_path) {
        // Load last saved values for this template, if any
        try {
          const saved = window.localStorage.getItem(`mcms:fieldValues:${def.template_path}`);
          if (saved) {
            const parsed = JSON.parse(saved);
            if (parsed && typeof parsed === 'object') {
              if (parsed.issue_date && !parsed.invoice_date) parsed.invoice_date = parsed.issue_date;
              setFieldValues(prev => ({ ...parsed, ...prev }));
            }
          }
        } catch (_) {}
      }
    } catch (_) {}
  }, []);

  // Removed openPlaceholders/savePlaceholders

  useEffect(() => { refreshClients(); refreshDocs(); loadInvoiceDefinition(); }, [refreshClients, refreshDocs, loadInvoiceDefinition]);
  useEffect(() => {
    (async () => {
      try {
        const settings = await window.api?.businessSettings?.();
        const rec = (Array.isArray(settings) ? settings : []).find(r => Number(r.id) === BUSINESS_ID) || null;
        setSavePath(rec?.save_path || '');
      } catch (_) {}
    })();
  }, []);

  // Show DB path on Templates tab (for due diligence: see where dev app's DB is)
  useEffect(() => {
    if (activeTab !== 'templates') return;
    try {
      const p = typeof window.api?.getDbPath === 'function' ? window.api.getDbPath() : '';
      setDbPath(p || '');
    } catch (_) {
      setDbPath('');
    }
  }, [activeTab]);

  // Auto-watch template file: when it changes on disk, copy to staging and show message
  useEffect(() => {
    if (!excelTemplatePath || !savePath || typeof window.api?.watchTemplateFile !== 'function') return;
    const result = window.api.watchTemplateFile({
      businessId: BUSINESS_ID,
      templatePath: excelTemplatePath,
      onChange: () => {
        setMessage('Staging updated from template change.');
        setTimeout(() => setMessage(''), 2500);
      }
    });
    if (!result?.ok) return;
    return () => { window.api?.unwatchTemplateFile?.(); };
  }, [excelTemplatePath, savePath]);

  // Prefill invoice number (next) when opening invoices tab or on first load
  useEffect(() => {
    const run = async () => {
      try {
        const max = await window.api?.getMaxInvoiceNumber?.(BUSINESS_ID);
        const next = (max != null ? Number(max) : 0) + 1;
        if (!invoiceNumTouched) setInvoiceNumber(String(next));
      } catch (_) {}
    };
    if (activeTab === 'invoices' && !invoiceNumTouched) run();
  }, [activeTab, invoiceNumTouched]);

  // Inline check: is invoice number taken? Debounced
  useEffect(() => {
    let timer = null;
    setInvoiceNumError('');
    setInvoiceNumTaken(false);
    if (!invoiceNumber || String(invoiceNumber).trim() === '') return; // empty means auto
    const val = Number(invoiceNumber);
    if (!Number.isInteger(val) || val < 1) {
      setInvoiceNumError('Enter a valid positive number');
      setInvoiceNumTaken(false);
      return;
    }
    setInvoiceNumChecking(true);
    timer = setTimeout(async () => {
      try {
        const exists = await window.api?.documentNumberExists?.(BUSINESS_ID, 'invoice', val);
        setInvoiceNumTaken(!!exists);
      } catch (err) {
        setInvoiceNumError(err?.message || 'Could not validate number');
      } finally {
        setInvoiceNumChecking(false);
      }
    }, 300);
    return () => { if (timer) clearTimeout(timer); };
  }, [invoiceNumber]);
  // No longer auto-prefilling template fields; Excel handles dates

  // Auto-select client if clientQuery exactly matches a client name (case-insensitive)
  useEffect(() => {
    const q = (clientQuery || '').trim().toLowerCase();
    if (!q) { setSelectedClient(null); return; }
    const match = clients.find(c => String(c.name || '').toLowerCase() === q) || null;
    setSelectedClient(match);
  }, [clientQuery, clients]);
  // No invoice date management in UI; Excel/workflow handles dating

  // Persist field values per template path
  useEffect(() => {
    if (!excelTemplatePath) return;
    try { window.localStorage.setItem(`mcms:fieldValues:${excelTemplatePath}`, JSON.stringify(fieldValues)); } catch (_) {}
  }, [excelTemplatePath, fieldValues]);

  // No template scanning in Excel-driven flow
  // Watch documents folder and auto-refresh
  useEffect(() => {
    const api = window.api;
    if (!api || !api.watchDocuments || !api.onDocumentsChange) return () => {};
    api.watchDocuments({ businessId: BUSINESS_ID }).catch(() => {});
    const unsub = api.onDocumentsChange(async (payload) => {
      try {
        if (!payload || payload.businessId !== BUSINESS_ID) return;
        // Safe mode: only refresh the log; manual Sync triggers importer
        refreshDocs();
      } catch (_) {}
    });
    return () => { try { unsub?.(); api.unwatchDocuments?.({ businessId: BUSINESS_ID }); } catch (_) {} };
  }, [refreshDocs]);

  // Load snapshot when opening invoice modal for edit
  useEffect(() => {
    if (!invoiceModalOpen || invoiceModalMode !== 'edit' || !editingDocument) return;
    (async () => {
      try {
        const doc = await window.api?.getDocumentById?.(editingDocument.document_id);
        if (doc?.invoice_snapshot) {
          const snap = JSON.parse(doc.invoice_snapshot);
          setClientQuery(snap.client_name || '');
          setInvoiceDate(snap.invoice_date || todayISO);
          setDueDate(snap.due_date != null ? snap.due_date : 'On receipt');
          setLineItems(Array.isArray(snap.line_items) && snap.line_items.length ? snap.line_items.map(it => {
            const desc = (it.description || '').toString().trim();
            const useDesc = desc.toLowerCase() === '(from existing invoice)' ? '' : (it.description || '');
            return { date: it.date || todayISO, description: useDesc, amount: it.amount ?? '' };
          }) : [{ date: todayISO, description: '', amount: '' }]);
          setTotalOverride(snap.total_override != null && snap.total_override !== '' ? String(snap.total_override) : '');
          setAmountReceived(snap.amount_received != null && snap.amount_received !== '' ? String(snap.amount_received) : '');
          setDiscountDescription(snap.discount_description || '');
          setDiscountAmount(snap.discount_amount != null && snap.discount_amount !== '' ? String(snap.discount_amount) : '');
          if (snap.invoice_number != null) { setInvoiceNumber(String(snap.invoice_number)); setInvoiceNumTouched(true); }
        } else {
          setClientQuery(doc?.client_name || '');
          setInvoiceDate(doc?.document_date ? doc.document_date.slice(0, 10) : todayISO);
          setDueDate(doc?.due_date || 'On receipt');
          const fallbackLine = { date: doc?.document_date ? doc.document_date.slice(0, 10) : todayISO, description: '', amount: doc?.total_amount ?? '' };
          try {
            const fromFile = doc?.file_path ? await window.api?.getInvoiceLineItemsFromFile?.(doc.file_path) : [];
            const loaded = Array.isArray(fromFile) && fromFile.length ? fromFile.map(it => ({ date: it.date || fallbackLine.date, description: it.description ?? '', amount: it.amount ?? '' })) : [fallbackLine];
            setLineItems(loaded);
          } catch (_) {
            setLineItems([fallbackLine]);
          }
          setTotalOverride('');
          setAmountReceived('');
          setDiscountDescription('');
          setDiscountAmount('');
          if (doc?.number != null) { setInvoiceNumber(String(doc.number)); setInvoiceNumTouched(true); }
        }
      } catch (_) {}
    })();
  }, [invoiceModalOpen, invoiceModalMode, editingDocument?.document_id]);

  return (
    <div style={{ minHeight: '100vh', background: '#f1f5f9', color: '#0f172a' }}>
      <style dangerouslySetInnerHTML={{ __html: '@keyframes mcms-spin { to { transform: rotate(360deg); } }' }} />
      <header style={{ background: '#fff', borderBottom: '1px solid #e2e8f0' }}>
        <div style={{ maxWidth: 1100, margin: '0 auto', padding: '16px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 600 }}>Motti Cohen Music Services</div>
            <div style={{ display: 'flex', gap: 12, marginTop: 6 }}>
              <button onClick={()=>setActiveTab('invoices')} style={{ fontSize: 12, padding: '4px 8px', borderRadius: 6, border: '1px solid #e2e8f0', background: activeTab==='invoices' ? '#eef2ff' : '#fff', color: activeTab==='invoices' ? '#3730a3' : '#475569' }}>Invoices</button>
              <button onClick={()=>{ setActiveTab('contacts'); if (!clients.length) refreshClients(); }} style={{ fontSize: 12, padding: '4px 8px', borderRadius: 6, border: '1px solid #e2e8f0', background: activeTab==='contacts' ? '#eef2ff' : '#fff', color: activeTab==='contacts' ? '#3730a3' : '#475569' }}>Contacts</button>
              <button onClick={()=>setActiveTab('templates')} style={{ fontSize: 12, padding: '4px 8px', borderRadius: 6, border: '1px solid #e2e8f0', background: activeTab==='templates' ? '#eef2ff' : '#fff', color: activeTab==='templates' ? '#3730a3' : '#475569' }}>Templates</button>
            </div>
          </div>
        </div>
      </header>
      <main style={{ maxWidth: 1100, margin: '0 auto', padding: 24 }}>
        {activeTab === 'invoices' ? (
        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ fontSize: 16, fontWeight: 600 }}>Invoices</div>
            <button
              onClick={() => {
                setInvoiceModalMode('new');
                setEditingDocument(null);
                setClientQuery('');
                setSelectedClient(null);
                setInvoiceDate(todayISO);
                setDueDate('On receipt');
                setLineItems([{ date: todayISO, description: '', amount: '' }]);
                setTotalOverride('');
                setAmountReceived('');
                setDiscountDescription('');
                setDiscountAmount('');
                setInvoiceNumTouched(false);
                setInvoiceModalOpen(true);
              }}
              style={{ fontSize: 14, padding: '10px 16px', border: '1px solid #4f46e5', borderRadius: 6, color: '#fff', background: '#4f46e5' }}
            >New invoice</button>
          </div>
        </section>
        ) : null}

        {activeTab === 'templates' ? (
        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 11, color: '#64748b', marginBottom: 12, fontFamily: 'monospace', wordBreak: 'break-all' }}>
            Database: {dbPath || '…'}
          </div>
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 12 }}>Template and save folder</div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, padding: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Template</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                <button
                  style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}
                  onClick={async () => {
                    try {
                      if (!savePath) {
                        setError('Please set the save folder first so the staging file can be updated.');
                        return;
                      }
                      const file = await window.api.chooseFile({ title: 'Select invoice template (xlsx)', filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }] });
                      if (!file) return;
                      await window.api.copyTemplateToStaging(BUSINESS_ID, file);
                      await window.api.saveDocumentDefinition(BUSINESS_ID, { key: 'invoice_balance', doc_type: 'invoice', label: 'Invoice – Balance', template_path: file, is_active: 1, is_locked: 0 });
                      setExcelTemplatePath(file);
                      setMessage('Template and staging file updated. On first invoice after an update, Excel may ask for access — grant it once.');
                      setTimeout(() => setMessage(''), 4000);
                    } catch (err) { setError(err?.message || 'Unable to set template'); }
                  }}
                >Set template…</button>
                {excelTemplatePath ? (
                  <>
                    <button onClick={()=>window.api?.openPath?.(excelTemplatePath)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}>Open</button>
                    <button onClick={()=>window.api?.showItemInFolder?.(excelTemplatePath)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}>Show in Finder</button>
                  </>
                ) : null}
              </div>
              <div style={{ fontSize: 12, color: excelTemplatePath ? '#64748b' : '#b91c1c', marginTop: 6, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{excelTemplatePath || 'No template set'}</div>
            </div>
            <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, padding: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Save folder</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                <button
                  onClick={async ()=>{
                    try {
                      if (!window.api || typeof window.api.chooseDirectory !== 'function') {
                        setError('App not ready — try again in a moment.');
                        return;
                      }
                      const dir = await window.api.chooseDirectory({ title: 'Choose invoice save folder' });
                      if (!dir) {
                        setMessage('No folder selected');
                        setTimeout(() => setMessage(''), 1500);
                        return;
                      }
                      if (typeof window.api.updateBusinessSettings !== 'function') {
                        setError('Cannot save settings.');
                        return;
                      }
                      await window.api.updateBusinessSettings(BUSINESS_ID, { save_path: dir });
                      setSavePath(dir);
                      setError('');
                      setMessage('Save folder updated');
                      setTimeout(() => setMessage(''), 1200);
                    } catch (err) {
                      setError(err?.message || 'Unable to set folder');
                    }
                  }}
                  style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}
                >Set save folder…</button>
                {savePath ? (
                  <button onClick={()=>window.api?.openPath?.(savePath)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}>Open</button>
                ) : null}
              </div>
              <div style={{ fontSize: 12, color: savePath ? '#64748b' : '#b91c1c', marginTop: 6, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{savePath || 'Not set'}</div>
            </div>
          </div>
          <p style={{ fontSize: 12, color: '#64748b', marginTop: 12, marginBottom: 0, maxWidth: 520 }}>
            The app watches the template file and copies it to a staging file in the save folder when you save changes. Excel always uses that same file to generate PDFs, so you only need to grant access once. After a template update, the first time you create an invoice Excel may ask for access to the staging file — grant it and you won&apos;t be asked again for that path.
          </p>
        </section>
        ) : null}

        {activeTab === 'contacts' ? (
          <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 16, fontWeight: 600 }}>Contacts</div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={()=>refreshClients()} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Refresh</button>
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '320px 1fr', gap: 16, marginTop: 12 }}>
              <div style={{ borderRight: '1px solid #e2e8f0', paddingRight: 12 }}>
                <div style={{ display: 'flex', gap: 6, marginBottom: 8 }}>
                  <input value={contactQuery} onChange={e=>setContactQuery(e.target.value)} onFocus={()=>setContactFocus(true)} onBlur={()=>setTimeout(()=>setContactFocus(false),120)} placeholder="Search contacts…" style={{ flex: 1, fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                  <button onClick={async ()=>{
                    const name = (contactQuery || '').trim();
                    if (!name) { setError('Enter a name'); return; }
                    try {
                      const existing = await window.api?.getClientByName?.(BUSINESS_ID, name);
                      if (existing) { setMessage('Client exists'); setTimeout(()=>setMessage(''), 1000); setContactSelectedId(existing.client_id); await loadClientDetails(existing.client_id); return; }
                    } catch(_){}
                    try {
                      const id = await window.api?.addClient?.({ business_id: BUSINESS_ID, name });
                      await refreshClients();
                      setContactQuery('');
                      await loadClientDetails(id);
                      setMessage('Client created'); setTimeout(()=>setMessage(''), 1000);
                    } catch (err) { setError(err?.message || 'Unable to create client'); }
                  }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>New</button>
                </div>
                <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                  <div style={{ maxHeight: 420, overflow: 'auto' }}>
                    {(() => {
                      const q = (contactQuery || '').trim().toLowerCase();
                      const list = clients
                        .filter(c => !c.business_id || c.business_id === BUSINESS_ID)
                        .map(c => {
                          if (!q) return { c, score: 0 };
                          const hay = `${c.name || ''}\n${c.email || ''}\n${c.phone || ''}`.toLowerCase();
                          const i = hay.indexOf(q);
                          const score = i < 0 ? Infinity : i + Math.abs(hay.length - q.length) * 0.01;
                          return { c, score };
                        })
                        .filter(x => x.score !== Infinity)
                        .sort((a,b) => a.score - b.score);
                      const finalList = q ? list.map(x=>x.c) : clients;
                      if (!finalList.length) return (<div style={{ padding: 8, color: '#64748b' }}>No contacts</div>);
                      return finalList.map(c => (
                        <div key={c.client_id} onClick={()=>loadClientDetails(c.client_id)} style={{ padding: 8, cursor: 'pointer', background: contactSelectedId===c.client_id ? '#eef2ff' : '#fff', borderBottom: '1px solid #e2e8f0' }}>
                          <div style={{ fontSize: 14, fontWeight: 500 }}>{c.name}</div>
                          {(c.email || c.phone) ? (<div style={{ fontSize: 12, color: '#64748b' }}>{[c.email, c.phone].filter(Boolean).join(' • ')}</div>) : null}
                        </div>
                      ));
                    })()}
                  </div>
                </div>
              </div>

              <div>
                {!contactSelectedId ? (
                  <div style={{ color: '#64748b' }}>Select a contact to edit</div>
                ) : (
                  <div style={{ display: 'grid', gap: 16 }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <div style={{ fontSize: 14, fontWeight: 600 }}>Edit contact</div>
                      <div style={{ display: 'flex', gap: 8 }}>
                        <button disabled={contactBusy} onClick={async ()=>{ if (!contactDetails.client) return; setContactBusy(true); try { await window.api?.deleteClient?.(contactDetails.client.client_id); setMessage('Deleted'); setTimeout(()=>setMessage(''), 1000); setContactSelectedId(null); setContactDetails({ client: null, emails: [], phones: [], addresses: [] }); await refreshClients(); } catch (err) { setError(err?.message || 'Unable to delete'); } finally { setContactBusy(false); } }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Delete</button>
                        <button disabled={contactBusy} onClick={async ()=>{ if (!contactDetails.client) return; setContactBusy(true); try { const det = await window.api?.getClientDetails?.(contactDetails.client.client_id); if (det && det.client) setContactDetails({ client: det.client, emails: det.emails||[], phones: det.phones||[], addresses: det.addresses||[] }); setMessage('Reverted'); setTimeout(()=>setMessage(''), 800); } catch (err) { setError(err?.message || 'Unable to reload'); } finally { setContactBusy(false); } }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}>Revert</button>
                        <button disabled={contactBusy} onClick={async ()=>{
                          if (!contactDetails.client) return;
                          setContactBusy(true);
                          try {
                            const draft = contactDetails;
                            const payload = {
                              name: draft.client?.name || '',
                              emails: (draft.emails||[]).map(e=>({ label: e.label||null, email: e.email||'', is_primary: e.is_primary?1:0 })),
                              phones: (draft.phones||[]).map(p=>({ label: p.label||null, phone: p.phone||'', is_primary: p.is_primary?1:0 })),
                              addresses: (draft.addresses||[]).map(a=>({ label: a.label||null, address1: a.address1||'', address2: a.address2||'', town: a.town||'', postcode: a.postcode||'', country: a.country||'', is_primary: a.is_primary?1:0 }))
                            };
                            await window.api?.saveClientDetails?.(draft.client.client_id, payload);
                            await refreshClients();
                            setMessage('Saved'); setTimeout(()=>setMessage(''), 1200);
                          } catch (err) { setError(err?.message || 'Unable to save'); }
                          finally { setContactBusy(false); }
                        }} style={{ fontSize: 12, padding: '6px 10px', borderRadius: 6, color: '#fff', background: contactBusy ? '#4f46e588' : '#4f46e5', border: 'none' }}>Save</button>
                      </div>
                    </div>

                    <div style={{ display: 'grid', gap: 12 }}>
                      <div style={{ display: 'flex', flexDirection: 'column' }}>
                        <label style={{ fontSize: 12, color: '#64748b' }}>Name</label>
                        <input value={contactDetails.client?.name || ''} onChange={e=>setContactDetails(prev=>({ ...prev, client: { ...(prev.client||{}), name: e.target.value } }))} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, maxWidth: 420 }} />
                      </div>

                      <div>
                        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Emails</div>
                        {(contactDetails.emails||[]).map((row, idx) => (
                          <div key={`em-${idx}`} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                            <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.emails = [...(prev.emails||[])]; next.emails[idx] = { ...next.emails[idx], label: e.target.value }; return next; })} style={{ width: 120, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="email@example.com" value={row.email||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.emails = [...(prev.emails||[])]; next.emails[idx] = { ...next.emails[idx], email: e.target.value }; return next; })} style={{ flex: 1, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="emailPrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const next={...prev}; next.emails = (prev.emails||[]).map((e,i)=>({ ...e, is_primary: i===idx?1:0 })); return next; })} /> Primary</label>
                            <button onClick={()=>setContactDetails(prev=>{ const next={...prev}; next.emails = (prev.emails||[]).filter((_,i)=>i!==idx); return next; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                          </div>
                        ))}
                        <button onClick={()=>setContactDetails(prev=>({ ...prev, emails: [...(prev.emails||[]), { label: '', email: '', is_primary: (prev.emails||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add email</button>
                      </div>

                      <div>
                        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Phones</div>
                        {(contactDetails.phones||[]).map((row, idx) => (
                          <div key={`ph-${idx}`} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                            <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.phones = [...(prev.phones||[])]; next.phones[idx] = { ...next.phones[idx], label: e.target.value }; return next; })} style={{ width: 120, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="Phone" value={row.phone||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.phones = [...(prev.phones||[])]; next.phones[idx] = { ...next.phones[idx], phone: e.target.value }; return next; })} style={{ flex: 1, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="phonePrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const next={...prev}; next.phones = (prev.phones||[]).map((p,i)=>({ ...p, is_primary: i===idx?1:0 })); return next; })} /> Primary</label>
                            <button onClick={()=>setContactDetails(prev=>{ const next={...prev}; next.phones = (prev.phones||[]).filter((_,i)=>i!==idx); return next; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                          </div>
                        ))}
                        <button onClick={()=>setContactDetails(prev=>({ ...prev, phones: [...(prev.phones||[]), { label: '', phone: '', is_primary: (prev.phones||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add phone</button>
                      </div>

                      <div>
                        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Addresses</div>
                        {(contactDetails.addresses||[]).map((row, idx) => (
                          <div key={`ad-${idx}`} style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 8, marginBottom: 8, border: '1px solid #e2e8f0', borderRadius: 8, padding: 8 }}>
                            <div style={{ gridColumn: 'span 2' }}>
                              <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.addresses = [...(prev.addresses||[])]; next.addresses[idx] = { ...next.addresses[idx], label: e.target.value }; return next; })} style={{ width: 180, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            </div>
                            <input placeholder="Address line 1" value={row.address1||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], address1: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="Address line 2" value={row.address2||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], address2: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="Town/City" value={row.town||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], town: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="Postcode" value={row.postcode||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], postcode: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <input placeholder="Country" value={row.country||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], country: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="addrPrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const next={...prev}; next.addresses = (prev.addresses||[]).map((a,i)=>({ ...a, is_primary: i===idx?1:0 })); return next; })} /> Primary</label>
                              <button onClick={()=>setContactDetails(prev=>{ const next={...prev}; next.addresses = (prev.addresses||[]).filter((_,i)=>i!==idx); return next; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                            </div>
                          </div>
                        ))}
                        <button onClick={()=>setContactDetails(prev=>({ ...prev, addresses: [...(prev.addresses||[]), { label: '', address1: '', address2: '', town: '', postcode: '', country: '', is_primary: (prev.addresses||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add address</button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </section>
        ) : null}

        {activeTab === 'invoices' ? (
        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <div style={{ fontSize: 16, fontWeight: 600 }}>Invoice Log</div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={refreshDocs} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Refresh</button>
              <button onClick={async ()=>{ try { await window.api?.indexInvoicesFromFilenames?.({ businessId: BUSINESS_ID }); setMessage('Synced from folder'); setTimeout(()=>setMessage(''), 800); refreshDocs(); } catch (err) { setError(err?.message || 'Sync failed'); } }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Sync from folder</button>
            </div>
          </div>
          {docsLoading ? (<div style={{ fontSize: 14, color: '#64748b' }}>Loading…</div>) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
              <thead>
                <tr style={{ background: '#f8fafc' }}>
                  <th onClick={()=>toggleSort('invoice')} role="button" style={{ userSelect: 'none', cursor: 'pointer', textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>
                    Invoice {sortKey==='invoice' ? (sortDir==='asc' ? '▲' : '▼') : ''}
                  </th>
                  <th onClick={()=>toggleSort('client')} role="button" style={{ userSelect: 'none', cursor: 'pointer', textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>
                    Client {sortKey==='client' ? (sortDir==='asc' ? '▲' : '▼') : ''}
                  </th>
                  <th onClick={()=>toggleSort('created')} role="button" style={{ userSelect: 'none', cursor: 'pointer', textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>
                    Created {sortKey==='created' ? (sortDir==='asc' ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ textAlign: 'right', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {sortedDocs.map(d => (
                  <tr key={d.document_id} style={{ borderTop: '1px solid #f1f5f9', background: (String(d.status||'').toLowerCase()==='paid') ? '#dcfce7' : '#fee2e2' }}>
                    <td style={{ padding: '8px' }}>Invoice #{d.number ?? ''}</td>
                    <td style={{ padding: '8px' }}>{d.display_client_name || d.client_name || ''}</td>
                    <td style={{ padding: '8px' }}>{d.event_date || d.document_date || d.created_at || ''}</td>
                    <td style={{ padding: '8px', textAlign: 'right' }}>
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <IconButton label="Open" onClick={() => window.api?.openPath?.(d.file_path)} size="sm"><EyeIcon /></IconButton>
                        <IconButton label="Reveal" onClick={() => window.api?.showItemInFolder?.(d.file_path)} size="sm"><RevealIcon /></IconButton>
                        <IconButton label="Edit" onClick={() => { setEditingDocument(d); setInvoiceModalMode('edit'); setInvoiceModalOpen(true); }} size="sm"><PencilIcon /></IconButton>
                        <IconButton
                          label="Delete"
                          onClick={async () => {
                          if (!d || !d.document_id) return;
                        try {
                          setDeleteModalData({
                            selected: d,
                            deletePdf: !!(d.file_path && (d.file_path || '').toLowerCase().endsWith('.pdf'))
                          });
                          setDeleteModalOpen(true);
                        } catch (err) { setError(err?.message || 'Unable to prepare delete'); }
                          }}
                        ><DeleteIcon /></IconButton>
                        {String(d.status || '').toLowerCase() === 'paid' ? (
                          <button style={{ fontSize: 12, padding: '6px 8px', marginLeft: 4, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={async ()=>{ try { await window.api?.updateDocumentStatus?.(d.document_id, { status: 'issued', paid_at: null }); setMessage('Marked unpaid'); setTimeout(()=>setMessage(''), 800); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to update'); } }}>Mark unpaid</button>
                        ) : (
                          <button style={{ fontSize: 12, padding: '6px 8px', marginLeft: 4, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={async ()=>{ try { await window.api?.updateDocumentStatus?.(d.document_id, { status: 'paid', paid_at: new Date().toISOString() }); setMessage('Marked paid'); setTimeout(()=>setMessage(''), 800); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to update'); } }}>Mark paid</button>
                        )}
                      </span>
                    </td>
                  </tr>
                ))}
                {!docs.length ? (
                  <tr><td colSpan="4" style={{ padding: '8px', color: '#64748b' }}>No invoices yet.</td></tr>
                ) : null}
              </tbody>
            </table>
          )}
        </section>
        ) : null}
      </main>
      {/* Invoice modal (New / Edit) */}
      {invoiceModalOpen ? (
        <div role="dialog" aria-modal="true" style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, zIndex: 50 }} onClick={(e)=>{ if(e.target===e.currentTarget) setInvoiceModalOpen(false); }}>
          <div style={{ width: 'min(640px, 96vw)', maxHeight: '90vh', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, boxShadow: '0 10px 30px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column' }}>
            <div style={{ padding: '14px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 16, fontWeight: 600 }}>{invoiceModalMode === 'edit' && editingDocument?.number != null ? `Edit invoice INV-${editingDocument.number}` : 'New invoice'}</div>
              <button onClick={()=>setInvoiceModalOpen(false)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Close</button>
            </div>
            <div style={{ padding: 16, overflow: 'auto', flex: 1 }}>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'flex-start' }}>
                <div style={{ display: 'flex', flexDirection: 'column', minWidth: 200, position: 'relative' }}>
                  <label style={{ fontSize: 12, color: '#64748b', marginBottom: 4, display: 'block' }}>Client</label>
                  <input value={clientQuery} onChange={e=>setClientQuery(e.target.value)} onFocus={()=>setClientFocus(true)} onBlur={()=>setTimeout(()=>setClientFocus(false), 120)} placeholder="Type a client name…" style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, boxSizing: 'border-box' }} />
                  {(clientFocus && (clientQuery||'').trim()) ? (
                    <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #e2e8f0', borderTop: 'none', borderRadius: '0 0 6px 6px', maxHeight: 180, overflow: 'auto', zIndex: 10 }}>
                      {clients.filter(c => !c.business_id || c.business_id === BUSINESS_ID).filter(c => String(c.name||'').toLowerCase().includes((clientQuery||'').trim().toLowerCase())).slice(0, 8).map(c => (
                        <div key={c.client_id} onMouseDown={()=>{ setSelectedClient(c); setClientQuery(c.name || ''); }} style={{ padding: 8, cursor: 'pointer' }}>{c.name}</div>
                      ))}
                    </div>
                  ) : null}
                  {clientQuery && (!selectedClient || String(selectedClient.name||'').toLowerCase() !== String(clientQuery||'').toLowerCase()) ? (
                    <button type="button" onClick={async ()=>{ const name = (clientQuery||'').trim(); if (!name) return; try { const existing = await window.api.getClientByName(BUSINESS_ID, name); if (existing) { setSelectedClient(existing); setClientQuery(existing.name||name); setMessage('Client exists'); setTimeout(()=>setMessage(''), 1000); return; } } catch(_){} const newId = await window.api.addClient({ business_id: BUSINESS_ID, name }); await refreshClients(); const row = await window.api.getClient(newId); if (row) setSelectedClient(row); setMessage('Saved'); setTimeout(()=>setMessage(''), 1000); }} style={{ fontSize: 12, padding: '4px 8px', marginTop: 4, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Save new client</button>
                  ) : null}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <label style={{ fontSize: 12, color: '#64748b', marginBottom: 4, display: 'block' }}>Invoice #</label>
                  <input type="number" min="1" value={invoiceNumber} onChange={e=>{ setInvoiceNumber(e.target.value); setInvoiceNumTouched(true); }} placeholder="auto" style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 100, boxSizing: 'border-box' }} />
                  <div style={{ minHeight: 20, marginTop: 2 }}>
                    {invoiceNumError ? <span style={{ fontSize: 12, color: '#b91c1c' }}>{invoiceNumError}</span> : null}
                    {invoiceNumber && !invoiceNumError && (invoiceNumTaken ? <span style={{ fontSize: 12, color: '#b91c1c' }}>Taken</span> : <span style={{ fontSize: 12, color: '#16a34a' }}>OK</span>)}
                  </div>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <label style={{ fontSize: 12, color: '#64748b', marginBottom: 4, display: 'block' }}>Invoice date</label>
                  <input type="date" value={invoiceDate} onChange={e=>setInvoiceDate(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, boxSizing: 'border-box' }} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', minWidth: 160 }}>
                  <label style={{ fontSize: 12, color: '#64748b', marginBottom: 4, display: 'block' }}>Due date</label>
                  <input value={dueDate} onChange={e=>setDueDate(e.target.value)} placeholder="On receipt or 30 days" style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, boxSizing: 'border-box' }} />
                </div>
              </div>
              <div style={{ marginTop: 12 }}>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Line items</div>
                {lineItems.map((item, idx) => (
                  <div key={idx} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                    <input type="date" value={item.date||''} onChange={e=>setLineItems(prev=>{ const n=[...prev]; n[idx]={ ...n[idx], date: e.target.value }; return n; })} style={{ width: 130, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                    <input placeholder="Description" value={item.description||''} onChange={e=>setLineItems(prev=>{ const n=[...prev]; n[idx]={ ...n[idx], description: e.target.value }; return n; })} style={{ flex: 1, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                    <input type="number" step="0.01" placeholder="Amount" value={item.amount===''?'':item.amount} onChange={e=>setLineItems(prev=>{ const n=[...prev]; n[idx]={ ...n[idx], amount: e.target.value }; return n; })} style={{ width: 90, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                    <button type="button" onClick={()=>setLineItems(prev=>prev.filter((_,i)=>i!==idx))} disabled={lineItems.length<=1} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                  </div>
                ))}
                <button type="button" onClick={()=>setLineItems(prev=>[...prev, { date: invoiceDate||todayISO, description: '', amount: '' }])} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add line</button>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, marginTop: 12 }}>
                <div><label style={{ fontSize: 12, color: '#64748b' }}>Total override</label><input type="number" step="0.01" value={totalOverride} onChange={e=>setTotalOverride(e.target.value)} placeholder="auto" style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 100, display: 'block' }} /></div>
                <div><label style={{ fontSize: 12, color: '#64748b' }}>Amount received</label><input type="number" step="0.01" value={amountReceived} onChange={e=>setAmountReceived(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 100, display: 'block' }} /></div>
                <div><label style={{ fontSize: 12, color: '#64748b' }}>Discount desc</label><input value={discountDescription} onChange={e=>setDiscountDescription(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 120, display: 'block' }} /></div>
                <div><label style={{ fontSize: 12, color: '#64748b' }}>Discount amt</label><input type="number" step="0.01" value={discountAmount} onChange={e=>setDiscountAmount(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 80, display: 'block' }} /></div>
              </div>
              <div style={{ marginTop: 16, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: 12 }}>
                {createBusy ? (
                  <>
                    <Spinner size={22} />
                    <span style={{ fontSize: 14, color: '#64748b' }}>
                      {invoiceModalMode === 'edit' ? 'Regenerating PDF…' : 'Generating PDF…'}
                    </span>
                  </>
                ) : null}
                <button
                  disabled={createBusy}
                  style={{ fontSize: 14, padding: '10px 16px', border: '1px solid #4f46e5', borderRadius: 6, color: '#fff', background: createBusy ? '#a5b4fc' : '#4f46e5', cursor: createBusy ? 'not-allowed' : 'pointer' }}
                  onClick={async ()=>{
                    setError('');
                    setCreateBusy(true);
                    try {
                      if (!savePath) { const dir = await window.api?.chooseDirectory?.({ title: 'Choose invoice save folder' }); if (!dir) return; await window.api?.updateBusinessSettings?.(BUSINESS_ID, { save_path: dir }); setSavePath(dir); }
                      const nameTyped = (clientQuery||'').trim();
                      if (!nameTyped) { setError('Enter a client name'); setCreateBusy(false); return; }
                      const validItems = lineItems.filter(it => (it.description||'').trim() || (it.amount !== '' && it.amount != null));
                      if (validItems.length === 0) { setError('Add at least one line item'); setCreateBusy(false); return; }
                      const invNum = invoiceModalMode === 'edit' && editingDocument?.number != null ? editingDocument.number : (invoiceNumber && Number.isFinite(Number(invoiceNumber)) ? Number(invoiceNumber) : undefined);
                      if (invoiceModalMode !== 'edit' && invNum != null && (invoiceNumTaken || invoiceNumError)) { setError('Invoice number invalid or taken'); setCreateBusy(false); return; }
                      let clientForInvoice = selectedClient;
                      if (!clientForInvoice && nameTyped) { try { const existing = await window.api?.getClientByName?.(BUSINESS_ID, nameTyped); if (existing) clientForInvoice = existing; } catch(_) {} if (!clientForInvoice) { const newId = await window.api?.addClient?.({ business_id: BUSINESS_ID, name: nameTyped }); await refreshClients(); const row = await window.api?.getClient?.(newId); if (row) clientForInvoice = row; } }
                      let clientOverride = { name: nameTyped };
                      if (clientForInvoice?.client_id) { const det = await window.api?.getClientDetails?.(clientForInvoice.client_id); if (det?.client) { const pe = (det.emails||[]).find(e=>e.is_primary) || (det.emails||[])[0]; const pp = (det.phones||[]).find(p=>p.is_primary) || (det.phones||[])[0]; const pa = (det.addresses||[]).find(a=>a.is_primary) || (det.addresses||[])[0]; clientOverride = { name: det.client.name || nameTyped, email: pe?.email || '', phone: pp?.phone || '', address1: pa?.address1 || '', address2: pa?.address2 || '', town: pa?.town || '', postcode: pa?.postcode || '' }; } }
                      const invDate = invoiceDate || todayISO;
                      let dueDateValue = (dueDate||'').trim();
                      const daysMatch = dueDateValue.match(/^(\d+)\s*days?$/i);
                      if (daysMatch) { const d = new Date(invDate); d.setDate(d.getDate() + parseInt(daysMatch[1], 10)); dueDateValue = d.toISOString().slice(0, 10); }
                      const items = validItems.map(it => ({ description: (it.description||'').trim(), amount: Number(it.amount)||0, date: it.date||invDate }));
                      const autoTotal = items.reduce((s,it)=>s+(Number.isFinite(it.amount)?it.amount:0), 0);
                      const totalVal = totalOverride !== '' && Number.isFinite(Number(totalOverride)) ? Number(totalOverride) : autoTotal;
                      const fieldValues = { invoice_date: invDate, due_date: dueDateValue || 'On receipt', amount_received: amountReceived !== '' && Number.isFinite(Number(amountReceived)) ? Number(amountReceived) : 0 };
                      if ((discountDescription||'').trim()) fieldValues.discount_description = discountDescription.trim();
                      if (discountAmount !== '' && Number(discountAmount) !== 0) fieldValues.discount_amount = Number(discountAmount);
                      const formSnapshot = JSON.stringify({ client_name: nameTyped, line_items: items.map(it => ({ date: it.date, description: it.description, amount: it.amount })), invoice_date: invDate, due_date: dueDateValue || 'On receipt', total_override: totalOverride, amount_received: amountReceived, discount_description: discountDescription, discount_amount: discountAmount, invoice_number: invNum });
                      if (invoiceModalMode === 'edit' && editingDocument?.number != null) {
                        const sameNumDocs = await window.api?.getDocumentsByNumber?.(BUSINESS_ID, 'invoice', editingDocument.number) || [];
                        for (const doc of sameNumDocs) { try { await window.api?.deleteDocument?.(doc.document_id, { removeFile: true }); } catch(_) {} }
                      }
                      const res = await window.api?.createMCMSInvoice?.({
                        business_id: BUSINESS_ID, definition_key: 'invoice_balance', client_override: clientOverride, line_items: items, document_date: invDate, due_date: dueDateValue || null, total_amount: totalVal,
                        invoice_number: invNum, amount_received: fieldValues.amount_received, discount_description: fieldValues.discount_description, discount_amount: fieldValues.discount_amount, field_values: fieldValues, form_snapshot: formSnapshot
                      });
                      if (res?.number != null) {
                        setMessage(invoiceModalMode === 'edit' ? `Invoice INV-${res.number} regenerated` : `Invoice INV-${res.number} created`);
                        setTimeout(() => setMessage(''), 2000);
                        if (invoiceModalMode !== 'edit') setInvoiceNumber(String(Number(res.number) + 1)); setInvoiceNumTouched(false);
                        setInvoiceModalOpen(false); setEditingDocument(null);
                        await refreshDocs();
                        if (res.file_path) { try { await window.api?.showItemInFolder?.(res.file_path); } catch(_) {} }
                      }
                    } catch (err) { setError(err?.message || 'Unable to create invoice'); } finally { setCreateBusy(false); }
                  }}
                >{invoiceModalMode === 'edit' ? 'Regenerate' : 'Create invoice'}</button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
      {/* Placeholders modal removed */}
      {contactModalOpen ? (
        <div role="dialog" aria-modal="true" style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }} onClick={(e)=>{ if(e.target===e.currentTarget) setContactModalOpen(false); }}>
          <div style={{ width: 'min(1000px, 96vw)', maxHeight: '90vh', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, boxShadow: '0 10px 30px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column' }}>
            <div style={{ padding: '14px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 16, fontWeight: 600 }}>Edit contact</div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button disabled={contactBusy || !contactDetails.client} onClick={async ()=>{ if (!contactDetails.client) return; setContactBusy(true); try { await window.api?.deleteClient?.(contactDetails.client.client_id); setMessage('Deleted'); setTimeout(()=>setMessage(''), 1000); setContactSelectedId(null); setContactDetails({ client: null, emails: [], phones: [], addresses: [] }); await refreshClients(); setContactModalOpen(false); } catch (err) { setError(err?.message || 'Unable to delete'); } finally { setContactBusy(false); } }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Delete</button>
                <button disabled={contactBusy || !contactDetails.client} onClick={async ()=>{ if (!contactDetails.client) return; setContactBusy(true); try { const det = await window.api?.getClientDetails?.(contactDetails.client.client_id); if (det && det.client) setContactDetails({ client: det.client, emails: det.emails||[], phones: det.phones||[], addresses: det.addresses||[] }); setMessage('Reverted'); setTimeout(()=>setMessage(''), 800); } catch (err) { setError(err?.message || 'Unable to reload'); } finally { setContactBusy(false); } }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}>Revert</button>
                <button onClick={()=>setContactModalOpen(false)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Close</button>
                <button disabled={contactBusy || !contactDetails.client} onClick={async ()=>{ if (!contactDetails.client) return; setContactBusy(true); try { const draft = contactDetails; const payload = { name: draft.client?.name || '', emails: (draft.emails||[]).map(e=>({ label: e.label||null, email: e.email||'', is_primary: e.is_primary?1:0 })), phones: (draft.phones||[]).map(p=>({ label: p.label||null, phone: p.phone||'', is_primary: p.is_primary?1:0 })), addresses: (draft.addresses||[]).map(a=>({ label: a.label||null, address1: a.address1||'', address2: a.address2||'', town: a.town||'', postcode: a.postcode||'', country: a.country||'', is_primary: a.is_primary?1:0 })) }; await window.api?.saveClientDetails?.(draft.client.client_id, payload); await refreshClients(); setMessage('Saved'); setTimeout(()=>setMessage(''), 1200); } catch (err) { setError(err?.message || 'Unable to save'); } finally { setContactBusy(false); } }} style={{ fontSize: 12, padding: '6px 10px', borderRadius: 6, color: '#fff', background: contactBusy ? '#4f46e588' : '#4f46e5', border: 'none' }}>Save</button>
              </div>
            </div>
            <div style={{ padding: 14, overflow: 'auto' }}>
              {!contactDetails.client ? (
                <div style={{ color: '#64748b' }}>Loading…</div>
              ) : (
                <div style={{ display: 'grid', gap: 16 }}>
                  <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <label style={{ fontSize: 12, color: '#64748b' }}>Name</label>
                    <input value={contactDetails.client?.name || ''} onChange={e=>setContactDetails(prev=>({ ...prev, client: { ...(prev.client||{}), name: e.target.value } }))} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, maxWidth: 420 }} />
                  </div>

                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Emails</div>
                    {(contactDetails.emails||[]).map((row, idx) => (
                      <div key={`mem-${idx}`} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                        <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.emails = [...(prev.emails||[])]; next.emails[idx] = { ...next.emails[idx], label: e.target.value }; return next; })} style={{ width: 120, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="email@example.com" value={row.email||''} onChange={e=>setContactDetails(prev=>{ const next={...prev}; next.emails = [...(prev.emails||[])]; next.emails[idx] = { ...next.emails[idx], email: e.target.value }; return next; })} style={{ flex: 1, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="modalEmailPrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const n={...prev}; n.emails = (prev.emails||[]).map((e,i)=>({ ...e, is_primary: i===idx?1:0 })); return n; })} /> Primary</label>
                        <button onClick={()=>setContactDetails(prev=>{ const n={...prev}; n.emails = (prev.emails||[]).filter((_,i)=>i!==idx); return n; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                      </div>
                    ))}
                    <button onClick={()=>setContactDetails(prev=>({ ...prev, emails: [...(prev.emails||[]), { label: '', email: '', is_primary: (prev.emails||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add email</button>
                  </div>

                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Phones</div>
                    {(contactDetails.phones||[]).map((row, idx) => (
                      <div key={`mph-${idx}`} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                        <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.phones=[...(prev.phones||[])]; n.phones[idx]={ ...n.phones[idx], label: e.target.value }; return n; })} style={{ width: 120, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="Phone" value={row.phone||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.phones=[...(prev.phones||[])]; n.phones[idx]={ ...n.phones[idx], phone: e.target.value }; return n; })} style={{ flex: 1, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="modalPhonePrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const n={...prev}; n.phones=(prev.phones||[]).map((p,i)=>({ ...p, is_primary: i===idx?1:0 })); return n; })} /> Primary</label>
                        <button onClick={()=>setContactDetails(prev=>{ const n={...prev}; n.phones=(prev.phones||[]).filter((_,i)=>i!==idx); return n; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                      </div>
                    ))}
                    <button onClick={()=>setContactDetails(prev=>({ ...prev, phones: [...(prev.phones||[]), { label: '', phone: '', is_primary: (prev.phones||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add phone</button>
                  </div>

                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Addresses</div>
                    {(contactDetails.addresses||[]).map((row, idx) => (
                      <div key={`mad-${idx}`} style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 8, marginBottom: 8, border: '1px solid #e2e8f0', borderRadius: 8, padding: 8 }}>
                        <div style={{ gridColumn: 'span 2' }}>
                          <input placeholder="Label" value={row.label||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], label: e.target.value }; return n; })} style={{ width: 180, fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        </div>
                        <input placeholder="Address line 1" value={row.address1||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], address1: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="Address line 2" value={row.address2||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], address2: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="Town/City" value={row.town||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], town: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="Postcode" value={row.postcode||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], postcode: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <input placeholder="Country" value={row.country||''} onChange={e=>setContactDetails(prev=>{ const n={...prev}; n.addresses=[...(prev.addresses||[])]; n.addresses[idx]={ ...n.addresses[idx], country: e.target.value }; return n; })} style={{ fontSize: 13, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}><input type="radio" name="modalAddrPrimary" checked={row.is_primary===1||row.is_primary===true} onChange={()=>setContactDetails(prev=>{ const n={...prev}; n.addresses = (prev.addresses||[]).map((a,i)=>({ ...a, is_primary: i===idx?1:0 })); return n; })} /> Primary</label>
                          <button onClick={()=>setContactDetails(prev=>{ const n={...prev}; n.addresses=(prev.addresses||[]).filter((_,i)=>i!==idx); return n; })} style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #fecaca', borderRadius: 6, color: '#b91c1c', background: '#fff' }}>Remove</button>
                        </div>
                      </div>
                    ))}
                    <button onClick={()=>setContactDetails(prev=>({ ...prev, addresses: [...(prev.addresses||[]), { label: '', address1: '', address2: '', town: '', postcode: '', country: '', is_primary: (prev.addresses||[]).length?0:1 }] }))} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add address</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      ) : null}
      {deleteModalOpen ? (
        <div role="dialog" aria-modal="true" style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }} onClick={(e)=>{ if(e.target===e.currentTarget) setDeleteModalOpen(false); }}>
          <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', width: 420, maxWidth: '96vw', padding: 16, boxShadow: '0 10px 30px rgba(0,0,0,0.15)' }}>
            <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>Delete invoice</div>
            <div style={{ fontSize: 13, color: '#475569', marginBottom: 16 }}>This will remove the invoice record and delete the PDF from disk. Continue?</div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
              <button onClick={()=>setDeleteModalOpen(false)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Cancel</button>
              <button
                onClick={async ()=>{
                  const d = deleteModalData.selected;
                  if (!d || !d.document_id) { setDeleteModalOpen(false); return; }

                  const deleteByPathSafe = async (absPath) => {
                    if (!absPath) return;
                    try { await window.api?.deleteDocumentByPath?.({ businessId: BUSINESS_ID, absolutePath: absPath }); }
                    catch (e) { throw new Error(e?.message || 'Unable to delete by path'); }
                  };

                  try {
                    if (d.file_path) {
                      try { await deleteByPathSafe(d.file_path); } catch (e) { setError(e?.message || 'Unable to delete PDF'); }
                    }
                    await window.api.deleteDocument(d.document_id, { removeFile: false });

                    try {
                      const max = await window.api?.getMaxInvoiceNumber?.(BUSINESS_ID);
                      const next = Number.isFinite(Number(max)) ? Number(max) : 0;
                      await window.api?.setLastInvoiceNumber?.(BUSINESS_ID, next);
                    } catch (_) {}

                    setDeleteModalOpen(false);
                    setMessage('Deleted'); setTimeout(()=>setMessage(''), 1000); refreshDocs();
                  } catch (err) {
                    setDeleteModalOpen(false);
                    setError(err?.message || 'Unable to delete');
                  }
                }}
                style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #fecaca', borderRadius: 6, background: '#fff', color: '#b91c1c' }}
              >Delete</button>
            </div>
          </div>
        </div>
      ) : null}
      {/* Fixed-position toasts so they don't cause layout jump */}
      <div style={{ position: 'fixed', bottom: 16, right: 16, zIndex: 9999, display: 'flex', flexDirection: 'column', gap: 8, pointerEvents: 'none' }}>
        {error ? (
          <div style={{ background: '#fee2e2', color: '#991b1b', border: '1px solid #fecaca', padding: '10px 14px', borderRadius: 8, boxShadow: '0 4px 12px rgba(0,0,0,0.15)', maxWidth: 360 }}>{error}</div>
        ) : null}
        {message ? (
          <div style={{ background: '#dcfce7', color: '#166534', border: '1px solid #bbf7d0', padding: '10px 14px', borderRadius: 8, boxShadow: '0 4px 12px rgba(0,0,0,0.15)', maxWidth: 360 }}>{message}</div>
        ) : null}
      </div>
    </div>
  );
}

const root = createRoot(document.getElementById('root'));
root.render(<App />);

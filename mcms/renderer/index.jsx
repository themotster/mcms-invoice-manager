import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';

function App() {
  const BUSINESS_ID = 1; // MCMS
  const [activeTab, setActiveTab] = useState('invoices'); // 'invoices' | 'contacts'
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
  const [deleteModalData, setDeleteModalData] = useState({
    selected: null,
    xlsxDoc: null,
    pdfDoc: null,
    removeSelectedFile: false,
    deleteXlsx: false,
    deletePdf: false
  });
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
      let filtered = list;
      try {
        filtered = await window.api?.filterDocumentsByExistingFiles?.(list, { includeMissing: false });
      } catch (_) {}
      setDocs(Array.isArray(filtered) ? filtered : list);
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

  return (
    <div style={{ minHeight: '100vh', background: '#f1f5f9', color: '#0f172a' }}>
      <header style={{ background: '#fff', borderBottom: '1px solid #e2e8f0' }}>
        <div style={{ maxWidth: 1100, margin: '0 auto', padding: '16px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 600 }}>Motti Cohen Music Services</div>
            <div style={{ display: 'flex', gap: 12, marginTop: 6 }}>
              <button onClick={()=>setActiveTab('invoices')} style={{ fontSize: 12, padding: '4px 8px', borderRadius: 6, border: '1px solid #e2e8f0', background: activeTab==='invoices' ? '#eef2ff' : '#fff', color: activeTab==='invoices' ? '#3730a3' : '#475569' }}>Invoices</button>
              <button onClick={()=>{ setActiveTab('contacts'); if (!clients.length) refreshClients(); }} style={{ fontSize: 12, padding: '4px 8px', borderRadius: 6, border: '1px solid #e2e8f0', background: activeTab==='contacts' ? '#eef2ff' : '#fff', color: activeTab==='contacts' ? '#3730a3' : '#475569' }}>Contacts</button>
            </div>
          </div>
        </div>
      </header>
      <main style={{ maxWidth: 1100, margin: '0 auto', padding: 24 }}>
        {error ? (<div style={{ background: '#fee2e2', color: '#991b1b', border: '1px solid #fecaca', padding: 10, borderRadius: 6, marginBottom: 12 }}>{error}</div>) : null}
        {message ? (<div style={{ background: '#dcfce7', color: '#166534', border: '1px solid #bbf7d0', padding: 10, borderRadius: 6, marginBottom: 12 }}>{message}</div>) : null}
        {activeTab === 'invoices' ? (
        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
          <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between' }}>
            <div style={{ fontSize: 16, fontWeight: 600 }}>Generate from Excel template</div>
          <div style={{ marginTop: 8 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, padding: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Template</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                  <button
                    style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}
                    onClick={async () => {
                      try {
                        const file = await window.api.chooseFile({ title: 'Select invoice template (xlsx)', filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }] });
                        if (!file) return;
                        await window.api.saveDocumentDefinition(BUSINESS_ID, { key: 'invoice_balance', doc_type: 'invoice', label: 'Invoice – Balance', template_path: file, is_active: 1, is_locked: 0 });
                        setExcelTemplatePath(file);
                        setMessage('Template set'); setTimeout(() => setMessage(''), 1200);
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
                        const dir = await window.api?.chooseDirectory?.({ title: 'Choose invoice save folder' });
                        if (!dir) return;
                        await window.api?.updateBusinessSettings?.(BUSINESS_ID, { save_path: dir });
                        setSavePath(dir);
                        setMessage('Save folder updated'); setTimeout(()=>setMessage(''), 1200);
                      } catch (err) { setError(err?.message || 'Unable to set folder'); }
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
          </div>
          </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'flex-end', marginTop: 12 }}>
              <div style={{ display: 'flex', flexDirection: 'column', minWidth: 280, position: 'relative' }}>
                <label style={{ fontSize: 12, color: '#64748b' }}>Client</label>
                <input
                  value={clientQuery}
                  onChange={e=>setClientQuery(e.target.value)}
                  onFocus={()=>setClientFocus(true)}
                  onBlur={()=>setTimeout(()=>setClientFocus(false), 120)}
                  placeholder="Type a client name…"
                  style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }}
                />
                {(clientFocus && (clientQuery||'').trim()) ? (
                  <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #e2e8f0', borderTop: 'none', borderBottomLeftRadius: 6, borderBottomRightRadius: 6, maxHeight: 220, overflow: 'auto', zIndex: 10 }}>
                    {(() => {
                      const q = (clientQuery || '').trim().toLowerCase();
                      const scored = clients
                        .filter(c => !c.business_id || c.business_id === BUSINESS_ID)
                        .map(c => {
                          const name = String(c.name || '');
                          const email = String(c.email || '');
                          const hay = `${name}\n${email}`.toLowerCase();
                          const idx = hay.indexOf(q);
                          const score = idx < 0 ? Infinity : idx + Math.abs(hay.length - q.length) * 0.01;
                          return { c, score };
                        })
                        .filter(x => x.score !== Infinity)
                        .sort((a,b) => a.score - b.score)
                        .slice(0, 8);
                      if (!scored.length) {
                        return (
                          <div style={{ padding: 8, color: '#64748b' }}>No matches</div>
                        );
                      }
                      return scored.map(({ c }) => (
                        <div key={c.client_id} onMouseDown={()=>{ setSelectedClient(c); setClientQuery(c.name || ''); }} style={{ padding: 8, cursor: 'pointer', display: 'flex', flexDirection: 'column' }}>
                          <div style={{ fontSize: 14 }}>{c.name}</div>
                          {(c.email || c.phone) ? (<div style={{ fontSize: 12, color: '#64748b' }}>{[c.email, c.phone].filter(Boolean).join(' • ')}</div>) : null}
                        </div>
                      ));
                    })()}
                  </div>
                ) : null}
                {clientQuery && (!selectedClient || (selectedClient && String(selectedClient.name || '').toLowerCase() !== String(clientQuery || '').toLowerCase())) ? (
                  <div style={{ marginTop: 6 }}>
                    <button
                      onClick={async ()=>{
                        try {
                          const name = (clientQuery || '').trim();
                          if (!name) return;
                          // Avoid duplicate by name (business scoped)
                          try {
                            const existing = await window.api.getClientByName(BUSINESS_ID, name);
                            if (existing) { setSelectedClient(existing); setClientQuery(existing.name || name); setMessage('Client exists; selected'); setTimeout(()=>setMessage(''), 1000); return; }
                          } catch (_) {}
                          const newId = await window.api.addClient({ business_id: BUSINESS_ID, name });
                          await refreshClients();
                          try { const row = await window.api.getClient(newId); if (row) setSelectedClient(row); } catch(_e) {}
                          // Open full editor modal for new client
                          await loadClientDetails(newId);
                          setContactModalOpen(true);
                          setMessage('Client saved'); setTimeout(()=>setMessage(''), 1000);
                          } catch (err) { setError(err?.message || 'Unable to save client'); }
                      }}
                      style={{ fontSize: 12, padding: '4px 8px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}
                    >Save new client</button>
                  </div>
                ) : null}
              </div>
              {/* Invoice date picker removed — Excel handles dates */}
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <label style={{ fontSize: 12, color: '#64748b' }}>Invoice #</label>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <input
                    type="number"
                    min="1"
                    value={invoiceNumber}
                    onChange={e=>{ setInvoiceNumber(e.target.value); setInvoiceNumTouched(true); }}
                    placeholder="auto"
                    style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 140 }}
                  />
                  <div style={{ fontSize: 12 }}>
                    {invoiceNumChecking ? (<span style={{ color: '#64748b' }}>Checking…</span>) : null}
                    {!invoiceNumChecking && invoiceNumError ? (<span style={{ color: '#b91c1c' }}>{invoiceNumError}</span>) : null}
                    {!invoiceNumChecking && !invoiceNumError && invoiceNumber && (invoiceNumTaken ? (<span style={{ color: '#b91c1c' }}>Number already in use</span>) : (<span style={{ color: '#16a34a' }}>Available</span>))}
                  </div>
                </div>
              </div>
              {/* Action button at far right */}
              <div style={{ marginLeft: 'auto' }}>
                <button
                  onClick={async ()=>{
                    try {
                      // Ensure save path
                      if (!savePath) {
                        const dir = await window.api?.chooseDirectory?.({ title: 'Choose invoice save folder' });
                        if (!dir) return; await window.api?.updateBusinessSettings?.(BUSINESS_ID, { save_path: dir }); setSavePath(dir);
                      }
                      const nameTyped = (clientQuery || '').trim();
                      if (!selectedClient && !nameTyped) { setError('Enter a client (name is used in filename)'); return; }
                      if (invoiceNumber && (invoiceNumTaken || invoiceNumError)) { setError('Invoice number invalid or taken'); return; }
                      const client = selectedClient || null;
                      const res = await window.api?.createNumberedWorkbookSimple?.({
                        business_id: BUSINESS_ID,
                        definition_key: 'invoice_balance',
                        invoice_number: (invoiceNumber && Number.isFinite(Number(invoiceNumber)) ? Number(invoiceNumber) : undefined),
                        client_name: client?.name || nameTyped
                      });
                      if (res && res.number != null) {
                        setMessage(`Workbook created for INV-${res.number}`); setTimeout(()=>setMessage(''), 1200);
                        try { const next = Number(res.number) + 1; if (Number.isFinite(next)) { setInvoiceNumber(String(next)); setInvoiceNumTouched(false); } } catch(_){}
                      }
                    } catch (err) { setError(err?.message || 'Unable to create workbook'); }
                  }}
                  style={{ fontSize: 12, padding: '8px 12px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#0f172a', background: '#f1f5f9' }}
                >New numbered workbook…</button>
              </div>
            
          </div>
          {/* Simplified: no additional fields or line items in Excel-driven flow */}
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
                      <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={() => window.api.openPath(d.file_path)}>Open</button>
                      <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={() => window.api.showItemInFolder(d.file_path)}>Reveal</button>
                      {String(d.status || '').toLowerCase() === 'paid' ? (
                        <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={async ()=>{ try { await window.api?.updateDocumentStatus?.(d.document_id, { status: 'issued', paid_at: null }); setMessage('Marked unpaid'); setTimeout(()=>setMessage(''), 800); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to update'); } }}>Mark unpaid</button>
                      ) : (
                        <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={async ()=>{ try { await window.api?.updateDocumentStatus?.(d.document_id, { status: 'paid', paid_at: new Date().toISOString() }); setMessage('Marked paid'); setTimeout(()=>setMessage(''), 800); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to update'); } }}>Mark paid</button>
                      )}
                      <button
                        style={{ fontSize: 12, padding: '6px 8px', border: '1px solid #fecaca', color: '#b91c1c', borderRadius: 6, background: '#fff' }}
                        onClick={async () => {
                          if (!d || !d.document_id) return;
                        try {
                          const same = await window.api?.getDocumentsByNumber?.(BUSINESS_ID, 'invoice', d.number);
                          const list = Array.isArray(same) ? same : [];
                          const lower = (p)=> (p||'').toString().toLowerCase();
                          const base = (p)=>{ const s=(p||'').toString(); const name = s.split(/\\\\|\//).pop() || ''; return name.replace(/\.[^.]+$/, ''); };
                          const dir = (p)=>{ const s=(p||'').toString(); const parts = s.split(/\\\\|\//); parts.pop(); return parts.join('/'); };
                          const selBase = base(d.file_path || '');
                          const selDir = dir(d.file_path || '');
                          // Only consider counterparts in the same directory and base
                          let xlsxDoc = list.find(x => dir(x.file_path||'') === selDir && base(x.file_path||'') === selBase && lower(x.file_path||'').endsWith('.xlsx')) || null;
                          let pdfDoc = list.find(x => dir(x.file_path||'') === selDir && base(x.file_path||'') === selBase && lower(x.file_path||'').endsWith('.pdf')) || null;
                          const isSelectedXlsx = !!(xlsxDoc && xlsxDoc.document_id === d.document_id);
                          const isSelectedPdf = !!(pdfDoc && pdfDoc.document_id === d.document_id);
                          setDeleteModalData({
                            selected: d,
                            xlsxDoc,
                            pdfDoc,
                            removeSelectedFile: false,
                            // Only pre-check counterpart files, not the selected one
                            deleteXlsx: !!(xlsxDoc && !isSelectedXlsx),
                            deletePdf: !!(pdfDoc && !isSelectedPdf)
                          });
                          setDeleteModalOpen(true);
                        } catch (err) { setError(err?.message || 'Unable to prepare delete'); }
                      }}
                      >Delete</button>
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
      </main>
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
            <div style={{ fontSize: 13, color: '#475569', marginBottom: 12 }}>Choose what to delete from disk. The database record is always removed.</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13 }}>
                <input type="checkbox" checked={deleteModalData.deleteXlsx} onChange={e=>setDeleteModalData(prev=>({ ...prev, deleteXlsx: e.target.checked }))} />
                Delete Excel workbook (.xlsx)
              </label>
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13 }}>
                <input type="checkbox" checked={deleteModalData.deletePdf} onChange={e=>setDeleteModalData(prev=>({ ...prev, deletePdf: e.target.checked }))} />
                Delete PDF (.pdf)
              </label>
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
              <button onClick={()=>setDeleteModalOpen(false)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Cancel</button>
              <button
                onClick={async ()=>{
                  const d = deleteModalData.selected;
                  if (!d || !d.document_id) { setDeleteModalOpen(false); return; }
                  const lower = (p)=> (p||'').toString().toLowerCase();
                  const xlsxDoc = deleteModalData.xlsxDoc || null;
                  const pdfDoc = deleteModalData.pdfDoc || null;
                  const { deleteXlsx, deletePdf } = deleteModalData;
                  const isSelectedXlsx = d.file_path && lower(d.file_path).endsWith('.xlsx');
                  const isSelectedPdf = d.file_path && lower(d.file_path).endsWith('.pdf');
                  const removeSelectedFile = (isSelectedXlsx && deleteXlsx) || (isSelectedPdf && deletePdf);

                  const deleteByPathSafe = async (absPath) => {
                    if (!absPath) return;
                    // Only allow deletion inside configured documents folder; never unlink outside
                    try { await window.api?.deleteDocumentByPath?.({ businessId: BUSINESS_ID, absolutePath: absPath }); }
                    catch (e) { throw new Error(e?.message || 'Unable to delete by path'); }
                  };

                  try {
                    // Remove selected document row; delete file by path if toggled
                    if (removeSelectedFile && d.file_path) {
                      try { await deleteByPathSafe(d.file_path); } catch (e) { setError(e?.message || 'Unable to delete selected file'); }
                    }
                    await window.api.deleteDocument(d.document_id, { removeFile: false });
                    if (deleteXlsx && isSelectedXlsx) {
                      try { await deleteByPathSafe(d.file_path); } catch (e) { setError(e?.message || 'Unable to delete Excel'); }
                    } else if (deleteXlsx && xlsxDoc && xlsxDoc.document_id !== d.document_id) {
                      try { await deleteByPathSafe(xlsxDoc.file_path); } catch (e) { setError(e?.message || 'Unable to delete Excel'); }
                      try { await window.api.deleteDocument(xlsxDoc.document_id, { removeFile: false }); } catch (_) {}
                    } else if (deleteXlsx && !xlsxDoc && isSelectedPdf && d.file_path) {
                      const twin = d.file_path.replace(/\.pdf$/i, '.xlsx');
                      try { await deleteByPathSafe(twin); } catch (_) {}
                    }

                    if (deletePdf && isSelectedPdf) {
                      try { await deleteByPathSafe(d.file_path); } catch (e) { setError(e?.message || 'Unable to delete PDF'); }
                    } else if (deletePdf && pdfDoc && pdfDoc.document_id !== d.document_id) {
                      try { await deleteByPathSafe(pdfDoc.file_path); } catch (e) { setError(e?.message || 'Unable to delete PDF'); }
                      try { await window.api.deleteDocument(pdfDoc.document_id, { removeFile: false }); } catch (_) {}
                    } else if (deletePdf && !pdfDoc && isSelectedXlsx && d.file_path) {
                      const twin = d.file_path.replace(/\.xlsx$/i, '.pdf');
                      try { await deleteByPathSafe(twin); } catch (_) {}
                    }

                    // Ensure invoice counter reflects current max after any deletions
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
    </div>
  );
}

const root = createRoot(document.getElementById('root'));
root.render(<App />);

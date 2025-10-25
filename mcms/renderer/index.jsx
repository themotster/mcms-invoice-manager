import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';

function App() {
  const BUSINESS_ID = 1; // MCMS
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');

  const [excelClientId, setExcelClientId] = useState('');
  const [excelDueDate, setExcelDueDate] = useState('');
  const [excelTemplatePath, setExcelTemplatePath] = useState('');
  const [excelBusy, setExcelBusy] = useState(false);
  const [items, setItems] = useState([]);
  const [phOpen, setPhOpen] = useState(false);
  const [phLoading, setPhLoading] = useState(false);
  const [phError, setPhError] = useState('');
  const [phList, setPhList] = useState([]); // [{ field_key, label, placeholder }]
  const [phEdits, setPhEdits] = useState(new Map()); // key -> placeholder

  const [docs, setDocs] = useState([]);
  const [docsLoading, setDocsLoading] = useState(true);

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
    } catch (_) {}
  }, []);

  const openPlaceholders = useCallback(async () => {
    setPhOpen(true);
    setPhLoading(true); setPhError('');
    try {
      const list = await window.api?.getMergeFields?.();
      const fields = Array.isArray(list) ? list : [];
      // Map to minimal info and include special tokens
      const base = fields.map(f => ({ field_key: f.field_key, label: f.label || f.field_key, placeholder: f.placeholder || '' }));
      // Add invoice_code and items as special entries
      const specials = [
        { field_key: 'invoice_code', label: 'Invoice code (e.g., INV-###)', placeholder: 'invoice_code', _special: true },
        { field_key: 'items', label: 'Line items anchor (table expand)', placeholder: 'items', _special: true }
      ];
      // Combine ensuring no duplicates
      const seen = new Set(base.map(x => x.field_key));
      const combined = base.slice();
      specials.forEach(s => { if (!seen.has(s.field_key)) combined.push(s); });
      setPhList(combined);
      setPhEdits(new Map());
    } catch (err) { setPhError(err?.message || 'Unable to load placeholders'); }
    finally { setPhLoading(false); }
  }, []);

  const savePlaceholders = useCallback(async () => {
    if (!phEdits.size) { setPhOpen(false); return; }
    setPhLoading(true); setPhError('');
    try {
      for (const [key, value] of phEdits.entries()) {
        const entry = phList.find(x => x.field_key === key);
        if (!entry) continue;
        // Skip specials
        if (entry._special) continue;
        await window.api?.saveMergeField?.({ field_key: key, label: entry.label || key, placeholder: String(value || '').trim() || null });
      }
      setPhEdits(new Map());
      await openPlaceholders();
    } catch (err) { setPhError(err?.message || 'Unable to save placeholders'); }
    finally { setPhLoading(false); }
  }, [phEdits, phList, openPlaceholders]);

  useEffect(() => { refreshClients(); refreshDocs(); loadInvoiceDefinition(); }, [refreshClients, refreshDocs, loadInvoiceDefinition]);
  // Watch documents folder and auto-refresh
  useEffect(() => {
    const api = window.api;
    if (!api || !api.watchDocuments || !api.onDocumentsChange) return () => {};
    api.watchDocuments({ businessId: BUSINESS_ID }).catch(() => {});
    const unsub = api.onDocumentsChange((payload) => {
      if (!payload || payload.businessId !== BUSINESS_ID) return;
      refreshDocs();
    });
    return () => { try { unsub?.(); api.unwatchDocuments?.({ businessId: BUSINESS_ID }); } catch (_) {} };
  }, [refreshDocs]);

  return (
    <div style={{ minHeight: '100vh', background: '#f1f5f9', color: '#0f172a' }}>
      <header style={{ background: '#fff', borderBottom: '1px solid #e2e8f0' }}>
        <div style={{ maxWidth: 1100, margin: '0 auto', padding: '16px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 600 }}>Motti Cohen Music Services</div>
            <div style={{ fontSize: 12, color: '#64748b' }}>Invoices</div>
          </div>
        </div>
      </header>
      <main style={{ maxWidth: 1100, margin: '0 auto', padding: 24 }}>
        {error ? (<div style={{ background: '#fee2e2', color: '#991b1b', border: '1px solid #fecaca', padding: 10, borderRadius: 6, marginBottom: 12 }}>{error}</div>) : null}
        {message ? (<div style={{ background: '#dcfce7', color: '#166534', border: '1px solid #bbf7d0', padding: 10, borderRadius: 6, marginBottom: 12 }}>{message}</div>) : null}

        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
          <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between' }}>
            <div style={{ fontSize: 16, fontWeight: 600 }}>Generate from Excel template</div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <button
                onClick={openPlaceholders}
                title="Show placeholders"
                style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, color: '#475569', background: '#fff' }}
              >Placeholders…</button>
              <div style={{ fontSize: 12, color: '#64748b', maxWidth: 520, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{excelTemplatePath ? `Template: ${excelTemplatePath}` : 'No template set'}</div>
            </div>
          </div>
          <div style={{ marginTop: 8 }}>
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
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'flex-end', marginTop: 12 }}>
            <div style={{ display: 'flex', flexDirection: 'column', minWidth: 220 }}>
              <label style={{ fontSize: 12, color: '#64748b' }}>Client</label>
              <select value={excelClientId} onChange={e => setExcelClientId(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }}>
                <option value="">Select…</option>
                {clients.map(c => (<option key={c.client_id} value={c.client_id}>{c.name}</option>))}
              </select>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <label style={{ fontSize: 12, color: '#64748b' }}>Due date</label>
              <input type="date" value={excelDueDate} onChange={e=>setExcelDueDate(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
            </div>
            <button
              disabled={excelBusy}
              onClick={async () => {
                setError(''); setExcelBusy(true);
                try {
                  const client = clients.find(c => String(c.client_id) === String(excelClientId));
                  if (!client) throw new Error('Select a client');
                  const computed = items.reduce((s, it) => { const q = Number(it.quantity), r = Number(it.rate); const ln = Number.isFinite(Number(it.amount)) ? Number(it.amount) : (Number.isFinite(q) && Number.isFinite(r) ? q*r : 0); return s + (Number.isFinite(ln) ? ln : 0); }, 0);
                  if (!items.length || !(computed > 0)) throw new Error('Add at least one line with a valid amount');
                  const res = await window.api.createMCMSInvoice({
                    business_id: BUSINESS_ID,
                    definition_key: 'invoice_balance',
                    client_override: {
                      name: client.name, email: client.email, phone: client.phone,
                      address1: client.address1 || client.address || '', address2: client.address2 || '', town: client.town || '', postcode: client.postcode || ''
                    },
                    line_items: items,
                    total_amount: computed,
                    due_date: excelDueDate || null
                  });
                  if (!res || !res.file_path) throw new Error('Invoice not created');
                  setMessage(`Invoice #${res?.number ?? ''} generated`); setTimeout(() => setMessage(''), 1500);
                  setItems([]); setExcelClientId(''); setExcelDueDate('');
                  refreshDocs();
                } catch (err) { setError(err?.message || 'Unable to generate invoice'); }
                finally { setExcelBusy(false); }
              }}
              style={{ fontSize: 12, padding: '8px 12px', borderRadius: 6, color: '#fff', background: excelBusy ? '#4f46e5AA' : '#4f46e5', border: 'none' }}
            >{excelBusy ? 'Generating…' : 'Generate Invoice'}</button>
          </div>
          <div style={{ marginTop: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 14, fontWeight: 600 }}>Line items</div>
              <div>
                <button onClick={() => setItems(prev => prev.concat([{ description: '', quantity: 1, unit: 'each', rate: 0 }]))} style={{ fontSize: 12, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Add line</button>
              </div>
            </div>
            <div style={{ overflowX: 'auto', border: '1px solid #e2e8f0', borderRadius: 8, marginTop: 8 }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
                <thead>
                  <tr style={{ background: '#f8fafc' }}>
                    <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #e2e8f0' }}>Description</th>
                    <th style={{ textAlign: 'right', padding: 8, borderBottom: '1px solid #e2e8f0', width: 80 }}>Qty</th>
                    <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #e2e8f0', width: 80 }}>Unit</th>
                    <th style={{ textAlign: 'right', padding: 8, borderBottom: '1px solid #e2e8f0', width: 120 }}>Rate</th>
                    <th style={{ textAlign: 'right', padding: 8, borderBottom: '1px solid #e2e8f0', width: 120 }}>Line Total</th>
                    <th style={{ padding: 8, borderBottom: '1px solid #e2e8f0', width: 80 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((it, idx) => {
                    const q = Number(it.quantity);
                    const r = Number(it.rate);
                    const line = Number.isFinite(Number(it.amount)) ? Number(it.amount) : (Number.isFinite(q) && Number.isFinite(r) ? q * r : 0);
                    return (
                      <tr key={`it-${idx}`} style={{ borderTop: '1px solid #f1f5f9' }}>
                        <td style={{ padding: 8 }}>
                          <input value={it.description || ''} onChange={e=>setItems(arr=>{ const next=[...arr]; next[idx] = { ...next[idx], description: e.target.value }; return next; })} style={{ width: '100%', fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} placeholder="Description" />
                        </td>
                        <td style={{ padding: 8, textAlign: 'right' }}>
                          <input type="number" step="0.01" value={Number.isFinite(q)?q:''} onChange={e=>setItems(arr=>{ const next=[...arr]; next[idx] = { ...next[idx], quantity: e.target.value }; return next; })} style={{ width: 90, fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, textAlign: 'right' }} />
                        </td>
                        <td style={{ padding: 8 }}>
                          <input value={it.unit || ''} onChange={e=>setItems(arr=>{ const next=[...arr]; next[idx] = { ...next[idx], unit: e.target.value }; return next; })} style={{ width: 80, fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6 }} />
                        </td>
                        <td style={{ padding: 8, textAlign: 'right' }}>
                          <input type="number" step="0.01" value={Number.isFinite(r)?r:''} onChange={e=>setItems(arr=>{ const next=[...arr]; next[idx] = { ...next[idx], rate: e.target.value }; return next; })} style={{ width: 120, fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, textAlign: 'right' }} />
                        </td>
                        <td style={{ padding: 8, textAlign: 'right' }}>£{Number.isFinite(line) ? line.toFixed(2) : '0.00'}</td>
                        <td style={{ padding: 8, textAlign: 'right' }}>
                          <button onClick={()=>setItems(arr=>arr.filter((_,i)=>i!==idx))} style={{ fontSize: 12, padding: '6px 8px', border: '1px solid #fecaca', color: '#b91c1c', borderRadius: 6, background: '#fff' }}>Remove</button>
                        </td>
                      </tr>
                    );
                  })}
                  {!items.length ? (
                    <tr><td colSpan="6" style={{ padding: 8, color: '#64748b' }}>No items yet. Add a line to begin.</td></tr>
                  ) : null}
                </tbody>
              </table>
            </div>
          </div>
        </section>

        <section style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>Invoice Log</div>
          {docsLoading ? (<div style={{ fontSize: 14, color: '#64748b' }}>Loading…</div>) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
              <thead>
                <tr style={{ background: '#f8fafc' }}>
                  <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>Invoice</th>
                  <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>Client</th>
                  <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>Created</th>
                  <th style={{ textAlign: 'right', padding: '8px', borderBottom: '1px solid #e2e8f0' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {docs.map(d => (
                  <tr key={d.document_id} style={{ borderTop: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '8px' }}>Invoice #{d.number ?? ''}</td>
                    <td style={{ padding: '8px' }}>{d.display_client_name || d.client_name || ''}</td>
                    <td style={{ padding: '8px' }}>{d.created_at || ''}</td>
                    <td style={{ padding: '8px', textAlign: 'right' }}>
                      <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={() => window.api.openPath(d.file_path)}>Open</button>
                      <button style={{ fontSize: 12, padding: '6px 8px', marginRight: 6, border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff' }} onClick={() => window.api.showItemInFolder(d.file_path)}>Reveal</button>
                      <button style={{ fontSize: 12, padding: '6px 8px', border: '1px solid #fecaca', color: '#b91c1c', borderRadius: 6, background: '#fff' }} onClick={async () => {
                        if (!d || !d.document_id) return;
                        const ok = window.confirm('Delete this record?');
                        if (!ok) return;
                        const removeFile = d.file_path ? window.confirm('Also delete the PDF file from disk?') : false;
                        try { await window.api.deleteDocument(d.document_id, { removeFile }); setMessage('Deleted'); setTimeout(()=>setMessage(''), 1000); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to delete'); }
                      }}>Delete</button>
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
      {phOpen ? (
        <div role="dialog" aria-modal="true" style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }} onClick={(e)=>{ if(e.target===e.currentTarget) setPhOpen(false); }}>
          <div style={{ width: 'min(900px, 95vw)', maxHeight: '85vh', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, boxShadow: '0 10px 30px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column' }}>
            <div style={{ padding: '14px 16px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 16, fontWeight: 600 }}>Placeholders</div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={async()=>{ const all = phList.map(f=>`{{${(f.placeholder||f.field_key)}}}`).join(', '); try { await window.api?.copyTextToClipboard?.(all); setMessage('Copied all'); setTimeout(()=>setMessage(''), 1000);} catch(_){} }} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Copy all</button>
                <button onClick={()=>setPhOpen(false)} style={{ fontSize: 12, padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Close</button>
                <button onClick={savePlaceholders} disabled={!phEdits.size || phLoading} style={{ fontSize: 12, padding: '6px 10px', borderRadius: 6, background: phEdits.size ? '#4f46e5' : '#4f46e588', color: '#fff', border: 'none' }}>{phLoading ? 'Saving…' : 'Save changes'}</button>
              </div>
            </div>
            {phError ? (<div style={{ padding: 12, color: '#b91c1c', background: '#fee2e2', borderBottom: '1px solid #fecaca' }}>{phError}</div>) : null}
            <div style={{ padding: 12, overflow: 'auto' }}>
              {phLoading ? (<div style={{ fontSize: 14, color: '#64748b' }}>Loading…</div>) : (
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
                  <thead>
                    <tr style={{ background: '#f8fafc' }}>
                      <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #e2e8f0' }}>Field</th>
                      <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #e2e8f0' }}>Placeholder</th>
                      <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #e2e8f0' }}>Token</th>
                      <th style={{ textAlign: 'right', padding: 8, borderBottom: '1px solid #e2e8f0' }}>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {phList.map(f => {
                      const editable = !f._special;
                      const current = phEdits.has(f.field_key) ? phEdits.get(f.field_key) : (f.placeholder || '');
                      const token = `{{${(current || f.field_key)}}}`;
                      return (
                        <tr key={f.field_key} style={{ borderTop: '1px solid #f1f5f9' }}>
                          <td style={{ padding: 8 }}>
                            <div style={{ fontWeight: 600 }}>{f.label || f.field_key}</div>
                            <div style={{ fontSize: 12, color: '#64748b' }}>{f.field_key}</div>
                          </td>
                          <td style={{ padding: 8 }}>
                            {editable ? (
                              <input value={current} onChange={e=>setPhEdits(prev=>{ const m=new Map(prev); m.set(f.field_key, e.target.value); return m; })} placeholder={f.field_key} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 260 }} />
                            ) : (
                              <div style={{ fontSize: 14, color: '#334155' }}>{current || f.field_key}</div>
                            )}
                          </td>
                          <td style={{ padding: 8 }}>
                            <code style={{ background: '#f8fafc', border: '1px solid #e2e8f0', padding: '2px 6px', borderRadius: 6 }}>{token}</code>
                          </td>
                          <td style={{ padding: 8, textAlign: 'right' }}>
                            <button onClick={async()=>{ try { await window.api?.copyTextToClipboard?.(token); setMessage('Copied'); setTimeout(()=>setMessage(''), 1000);} catch(_){} }} style={{ fontSize: 12, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, background: '#fff', color: '#475569' }}>Copy</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

const root = createRoot(document.getElementById('root'));
root.render(<App />);

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';

function App() {
  const BUSINESS_ID = 1; // MCMS
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');

  const [excelClientId, setExcelClientId] = useState('');
  const [excelAmount, setExcelAmount] = useState('');
  const [excelDueDate, setExcelDueDate] = useState('');
  const [excelTemplatePath, setExcelTemplatePath] = useState('');
  const [excelBusy, setExcelBusy] = useState(false);

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
      const items = await window.api?.getDocuments?.({ businessId: BUSINESS_ID, docType: 'invoice' });
      setDocs(Array.isArray(items) ? items : []);
    } catch (err) {
      setError(err?.message || 'Unable to load invoices');
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

  useEffect(() => { refreshClients(); refreshDocs(); loadInvoiceDefinition(); }, [refreshClients, refreshDocs, loadInvoiceDefinition]);

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
            <div style={{ fontSize: 12, color: '#64748b' }}>{excelTemplatePath ? `Template: ${excelTemplatePath}` : 'No template set'}</div>
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
              <label style={{ fontSize: 12, color: '#64748b' }}>Amount</label>
              <input type="number" step="0.01" value={excelAmount} onChange={e=>setExcelAmount(e.target.value)} style={{ fontSize: 14, padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: 6, width: 140 }} placeholder="0.00" />
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
                  const amt = Number(excelAmount);
                  if (!Number.isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
                  const res = await window.api.createNumberedDocument({
                    business_id: BUSINESS_ID,
                    doc_type: 'invoice',
                    definition_key: 'invoice_balance',
                    client_override: {
                      name: client.name, email: client.email, phone: client.phone,
                      address1: client.address1 || client.address || '', address2: client.address2 || '', town: client.town || '', postcode: client.postcode || ''
                    },
                    total_amount: amt,
                    due_date: excelDueDate || null
                  });
                  if (!res || !res.file_path) throw new Error('Invoice not created');
                  setMessage(`Invoice #${res?.number ?? ''} generated`); setTimeout(() => setMessage(''), 1500);
                  setExcelAmount(''); setExcelClientId(''); setExcelDueDate('');
                  refreshDocs();
                } catch (err) { setError(err?.message || 'Unable to generate invoice'); }
                finally { setExcelBusy(false); }
              }}
              style={{ fontSize: 12, padding: '8px 12px', borderRadius: 6, color: '#fff', background: excelBusy ? '#4f46e5AA' : '#4f46e5', border: 'none' }}
            >{excelBusy ? 'Generating…' : 'Generate Invoice'}</button>
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
                        const ok = window.confirm('Delete this record? This removes it from the list.');
                        if (!ok) return;
                        try { await window.api.deleteDocument(d.document_id, { removeFile: false }); setMessage('Deleted'); setTimeout(()=>setMessage(''), 1000); refreshDocs(); } catch (err) { setError(err?.message || 'Unable to delete'); }
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
    </div>
  );
}

const root = createRoot(document.getElementById('root'));
root.render(<App />);


import React, { useMemo, useState } from 'react';

function InvoiceCanvasEditor({ business, clients = [], onGenerate }) {
  const todayIso = new Date().toISOString().slice(0,10);
  const [clientId, setClientId] = useState('');
  const [billTo, setBillTo] = useState({ name: '', address1: '', address2: '', town: '', postcode: '' });
  const [issueDate, setIssueDate] = useState(todayIso);
  const [dueDate, setDueDate] = useState('');
  const [items, setItems] = useState([{ description: 'Performance fee', quantity: 1, unit: 'each', rate: 0 }]);
  const total = useMemo(() => items.reduce((s, it) => {
    const q = Number(it.quantity);
    const r = Number(it.rate);
    const amt = Number.isFinite(Number(it.amount)) ? Number(it.amount) : (Number.isFinite(q) && Number.isFinite(r) ? q * r : 0);
    return s + (Number.isFinite(amt) ? amt : 0);
  }, 0), [items]);

  const selectClient = (id) => {
    setClientId(id);
    const c = clients.find(x => String(x.client_id) === String(id));
    if (!c) return;
    setBillTo({
      name: c.name || '',
      address1: c.address1 || c.address || '',
      address2: c.address2 || '',
      town: c.town || '',
      postcode: c.postcode || ''
    });
  };

  const addRow = (type) => {
    const row = type === 'studio'
      ? { description: 'Studio time', quantity: 1, unit: 'hours', rate: 0 }
      : type === 'expense'
        ? { description: 'Expense', quantity: 1, unit: 'item', rate: 0 }
        : { description: 'Performance fee', quantity: 1, unit: 'each', rate: 0 };
    setItems(arr => arr.concat([row]));
  };

  const buildHtml = () => {
    const esc = (s) => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const fmt = (n) => {
      const v = Number(n);
      if (!Number.isFinite(v)) return '';
      try { return new Intl.NumberFormat('en-GB',{style:'currency',currency:'GBP'}).format(v);}catch(_){return `£${v.toFixed(2)}`;}
    };
    const rows = items.map(it => {
      const q = it.quantity != null && Number.isFinite(Number(it.quantity)) ? Number(it.quantity) : '';
      const r = it.rate != null && Number.isFinite(Number(it.rate)) ? fmt(it.rate) : '';
      const a = it.amount != null && Number.isFinite(Number(it.amount)) ? fmt(it.amount) : (Number.isFinite(Number(it.quantity)) && Number.isFinite(Number(it.rate)) ? fmt(Number(it.quantity) * Number(it.rate)) : '');
      return `<tr><td style="padding:8px;border-bottom:1px solid #e5e7eb">${esc(it.description||'')}</td><td style="padding:8px;border-bottom:1px solid #e5e7eb;text-align:right;width:80px">${q}</td><td style=\"padding:8px;border-bottom:1px solid #e5e7eb;width:80px\">${esc(it.unit||'')}</td><td style=\"padding:8px;border-bottom:1px solid #e5e7eb;text-align:right;width:120px\">${r}</td><td style=\"padding:8px;border-bottom:1px solid #e5e7eb;text-align:right;width:120px\">${a}</td></tr>`;
    }).join('');
    const addr = [billTo.address1, billTo.address2, billTo.town, billTo.postcode].filter(Boolean).map(esc).join('<br>');
    return `<!doctype html><html><head><meta charset="utf-8" /><title>Invoice</title><style>
      body{font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#111827;margin:40px;}
      h1{font-size:22px;margin:0 0 4px 0;} .muted{color:#6b7280;} table{width:100%;border-collapse:collapse;margin-top:12px}
    </style></head><body>
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <div><h1>${esc(business?.business_name || 'MCMS')}</h1><div class="muted">Invoice</div></div>
        <div style="text-align:right"><div style="font-size:28px;font-weight:700">Invoice</div><div class="muted">{{invoice_code}}</div></div>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-top:12px">
        <div><div style="font-weight:600">Bill to</div><div>${esc(billTo.name||'')}</div><div class="muted">${addr}</div></div>
        <div style="text-align:right"><div><span class="muted">Date:</span> ${esc(issueDate)}</div><div><span class="muted">Due:</span> ${esc(dueDate)}</div></div>
      </div>
      <table><thead><tr style="background:#f9fafb"><th style="text-align:left;padding:8px;border-bottom:1px solid #e5e7eb">Description</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:80px">Qty</th><th style="text-align:left;padding:8px;border-bottom:1px solid #e5e7eb;width:80px">Unit</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:120px">Rate</th><th style="text-align:right;padding:8px;border-bottom:1px solid #e5e7eb;width:120px">Line Total</th></tr></thead><tbody>${rows}</tbody></table>
      <div style="margin-top:12px;text-align:right;font-weight:700">Total: ${fmt(total)}</div>
    </body></html>`;
  };

  const handleGenerate = async () => {
    const html = buildHtml();
    const client = clients.find(x => String(x.client_id) === String(clientId));
    const client_override = {
      name: billTo.name,
      email: client?.email || '',
      phone: client?.phone || '',
      address1: billTo.address1, address2: billTo.address2, town: billTo.town, postcode: billTo.postcode
    };
    const payload = { inline_html: html, client_override, due_date: dueDate || null, line_items: items, total_amount: total };
    if (typeof onGenerate === 'function') onGenerate(payload);
  };

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-end gap-2">
        <div className="flex flex-col min-w-[220px]">
          <label className="text-xs text-slate-500">Client</label>
          <select value={clientId} onChange={e=>selectClient(e.target.value)} className="border rounded px-2 py-1 text-sm">
            <option value="">Select…</option>
            {clients.map(c => (<option key={c.client_id} value={c.client_id}>{c.name}</option>))}
          </select>
        </div>
        <div className="flex flex-col">
          <label className="text-xs text-slate-500">Issue date</label>
          <input type="date" value={issueDate} onChange={e=>setIssueDate(e.target.value)} className="border rounded px-2 py-1 text-sm" />
        </div>
        <div className="flex flex-col">
          <label className="text-xs text-slate-500">Due date</label>
          <input type="date" value={dueDate} onChange={e=>setDueDate(e.target.value)} className="border rounded px-2 py-1 text-sm" />
        </div>
        <button className="text-xs px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>addRow('gig')}>Add gig fee</button>
        <button className="text-xs px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>addRow('studio')}>Add studio time</button>
        <button className="text-xs px-3 py-2 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={()=>addRow('expense')}>Add expense</button>
        <div className="grow" />
        <button className="text-xs px-3 py-2 rounded bg-indigo-600 text-white hover:bg-indigo-500" onClick={handleGenerate}>Generate PDF</button>
      </div>

      <div className="border rounded shadow-sm bg-white" style={{ maxWidth: 860 }}>
        <div className="p-6">
          <div className="flex items-start justify-between">
            <div>
              <div className="text-xl font-bold text-slate-800">{business?.business_name || 'MCMS'}</div>
              <div className="text-xs text-slate-500">Invoice</div>
            </div>
            <div className="text-right">
              <div className="text-2xl font-bold text-slate-800">Invoice</div>
              <div className="text-xs text-slate-500">INV-—</div>
            </div>
          </div>
          <div className="mt-4 flex items-start justify-between">
            <div>
              <div className="font-medium text-slate-700">Bill to</div>
              <input className="block w-64 text-sm border-b" placeholder="Name" value={billTo.name} onChange={e=>setBillTo(v=>({...v,name:e.target.value}))} />
              <input className="block w-64 text-sm border-b" placeholder="Address line 1" value={billTo.address1} onChange={e=>setBillTo(v=>({...v,address1:e.target.value}))} />
              <input className="block w-64 text-sm border-b" placeholder="Address line 2" value={billTo.address2} onChange={e=>setBillTo(v=>({...v,address2:e.target.value}))} />
              <input className="block w-64 text-sm border-b" placeholder="Town/City" value={billTo.town} onChange={e=>setBillTo(v=>({...v,town:e.target.value}))} />
              <input className="block w-64 text-sm border-b" placeholder="Postcode" value={billTo.postcode} onChange={e=>setBillTo(v=>({...v,postcode:e.target.value}))} />
            </div>
            <div className="text-sm text-slate-700">
              <div><span className="text-slate-500">Date:</span> {issueDate}</div>
              <div><span className="text-slate-500">Due:</span> {dueDate || '—'}</div>
            </div>
          </div>
          <div className="mt-4 border rounded overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-700">
                  <th className="px-2 py-1 text-left">Description</th>
                  <th className="px-2 py-1 text-right" style={{width:'90px'}}>Qty/Hrs</th>
                  <th className="px-2 py-1 text-left" style={{width:'80px'}}>Unit</th>
                  <th className="px-2 py-1 text-right" style={{width:'120px'}}>Rate</th>
                  <th className="px-2 py-1 text-right" style={{width:'120px'}}>Line total</th>
                  <th className="px-2 py-1"></th>
                </tr>
              </thead>
              <tbody>
                {items.map((it, idx) => {
                  const q = Number(it.quantity);
                  const r = Number(it.rate);
                  const line = Number.isFinite(Number(it.amount)) ? Number(it.amount) : (Number.isFinite(q) && Number.isFinite(r) ? q*r : 0);
                  return (
                    <tr key={`row-${idx}`} className="border-t">
                      <td className="px-2 py-1"><input className="w-full border-b" value={it.description||''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = {...next[idx], description:e.target.value}; return next;})} placeholder="Description" /></td>
                      <td className="px-2 py-1 text-right"><input type="number" step="0.01" className="w-20 border-b text-right" value={Number.isFinite(q)?q:''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = {...next[idx], quantity:e.target.value}; return next;})} /></td>
                      <td className="px-2 py-1"><input className="w-20 border-b" value={it.unit||''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = {...next[idx], unit:e.target.value}; return next;})} /></td>
                      <td className="px-2 py-1 text-right"><input type="number" step="0.01" className="w-28 border-b text-right" value={Number.isFinite(r)?r:''} onChange={e=>setItems(arr=>{const next=[...arr]; next[idx] = {...next[idx], rate:e.target.value}; return next;})} /></td>
                      <td className="px-2 py-1 text-right">£{Number.isFinite(line)?line.toFixed(2):'0.00'}</td>
                      <td className="px-2 py-1 text-right"><button className="text-xs px-2 py-1 border rounded border-red-200 text-red-600 hover:bg-red-50" onClick={()=>setItems(arr=>arr.filter((_,i)=>i!==idx))}>Remove</button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="mt-2 text-right text-sm font-semibold">Total: £{total.toFixed(2)}</div>
        </div>
      </div>
    </div>
  );
}

export default InvoiceCanvasEditor;

export function ExcelTemplateEditor({ initialPath = '', onSaved }) {
  const [filePath, setFilePath] = useState(initialPath || '');
  const [sheet, setSheet] = useState('');
  const [sheets, setSheets] = useState([]);
  const [cells, setCells] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [changed, setChanged] = useState(new Map()); // key: `${r},${c}` -> value

  const load = async (p = filePath, s = '') => {
    if (!p) return;
    setLoading(true); setError('');
    try {
      const res = await window.api.readExcelSnapshot({ filePath: p, sheetName: s || undefined, maxRows: 40, maxCols: 20 });
      setSheets(Array.isArray(res?.sheets) ? res.sheets : []);
      setSheet(res?.sheet || '');
      setCells(Array.isArray(res?.cells) ? res.cells : []);
      setChanged(new Map());
    } catch (err) { setError(err?.message || 'Unable to read workbook'); }
    finally { setLoading(false); }
  };

  const pickFile = async () => {
    try { const f = await window.api.chooseFile({ title: 'Select Excel workbook', filters: [{ name: 'Excel', extensions: ['xlsx'] }] }); if (f) { setFilePath(f); load(f); } } catch (err) { setError(err?.message || 'Unable to choose file'); }
  };

  const onChangeCell = (r, c, v) => {
    setCells(prev => {
      const next = prev.map(row => row.slice());
      if (next[r] && c < next[r].length) next[r][c] = v;
      return next;
    });
    setChanged(prev => { const key = `${r+1},${c+1}`; const map = new Map(prev); map.set(key, v); return map; });
  };

  const save = async () => {
    if (!filePath || !sheet || !changed.size) return;
    try {
      const changes = Array.from(changed.entries()).map(([key, value]) => { const [row, col] = key.split(',').map(n => Number(n)); return { row, col, value }; });
      await window.api.writeExcelCells({ filePath, sheetName: sheet, changes });
      setChanged(new Map());
      if (typeof onSaved === 'function') onSaved({ filePath, sheet });
    } catch (err) { setError(err?.message || 'Unable to save changes'); }
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center gap-2">
        <button className="text-xs px-2 py-1 border rounded border-slate-300 text-slate-600 hover:bg-slate-50" onClick={pickFile}>{filePath ? 'Change file…' : 'Choose file…'}</button>
        {filePath ? <span className="text-xs text-slate-500 truncate" title={filePath}>{filePath}</span> : null}
        <div className="grow" />
        {sheets.length ? (
          <select value={sheet} onChange={e=>load(filePath, e.target.value)} className="text-xs border rounded px-2 py-1">
            {sheets.map(n => (<option key={n} value={n}>{n}</option>))}
          </select>
        ) : null}
        <button disabled={!changed.size} className="text-xs px-2 py-1 rounded bg-indigo-600 text-white hover:bg-indigo-500 disabled:opacity-50" onClick={save}>Save changes</button>
      </div>
      {error ? <div className="text-xs text-red-600">{error}</div> : null}
      {loading ? <div className="text-sm text-slate-500">Loading…</div> : null}
      {!loading && cells && cells.length ? (
        <div className="overflow-auto border rounded" style={{ maxHeight: 400 }}>
          <table className="text-sm">
            <thead>
              <tr>
                <th className="px-2 py-1 bg-slate-50 border-r border-b"></th>
                {Array.from({ length: cells[0].length }).map((_, c) => (
                  <th key={`h-${c}`} className="px-2 py-1 bg-slate-50 border-r border-b">{String.fromCharCode(65 + c)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cells.map((row, r) => (
                <tr key={`r-${r}`}>
                  <td className="px-2 py-1 bg-slate-50 border-r border-b text-right w-10">{r+1}</td>
                  {row.map((val, c) => (
                    <td key={`c-${r}-${c}`} className="px-1 py-0.5 border-b border-r">
                      <input className="w-40 px-1 py-0.5 text-xs border rounded" value={val ?? ''} onChange={e=>onChangeCell(r, c, e.target.value)} />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </div>
  );
}

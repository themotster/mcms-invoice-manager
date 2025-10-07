import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import ToastOverlay from './ToastOverlay';

const TOKEN_REGEX = /{{\s*([a-zA-Z0-9_.-]+)(?:\|([^}]+))?\s*}}/g;
const TOKEN_CHIP_CLASS = 'mail-token-chip';
const TOKEN_CHIP_STYLE = 'background:#eef2ff;border:1px solid #c7d2fe;border-radius:4px;padding:0 4px;margin:0 2px;display:inline-block;color:#312e81;font-size:0.85em;line-height:1.4;';
const TOKEN_FALLBACKS = {
  client_first_name: 'there'
};

const TOKEN_OPTIONS = [
  { key: 'client_name', label: 'Client name' },
  { key: 'client_first_name', label: 'Client first name' },
  { key: 'client_email', label: 'Client email' },
  { key: 'event_type', label: 'Event type' },
  { key: 'event_date', label: 'Event date' },
  { key: 'event_start', label: 'Event start' },
  { key: 'event_end', label: 'Event end' },
  { key: 'venue_name', label: 'Venue name' },
  { key: 'venue_town', label: 'Venue town' },
  { key: 'venue_postcode', label: 'Venue postcode' },
  { key: 'balance_amount', label: 'Balance amount' },
  { key: 'balance_due_date', label: 'Balance due' },
  { key: 'today', label: 'Today' }
];

const escapeHtml = (value) => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

const convertPlainTextToHtml = (text) => {
  const trimmed = String(text ?? '').replace(/\r/g, '');
  if (!trimmed.trim()) return '';
  const paragraphs = trimmed.split(/\n{2,}/).map(p => p.trim()).filter(Boolean);
  if (paragraphs.length) {
    return paragraphs
      .map(p => p.replace(/\n/g, '<br>'))
      .map(p => `<p style="margin:0 0 12px 0;">${p}</p>`)
      .join('');
  }
  return trimmed.replace(/\n/g, '<br>');
};

const normalizeTemplateBody = (input) => {
  if (input == null) return '';
  const str = String(input);
  if (!str.trim()) return '';
  if (/<[a-z][\s\S]*>/i.test(str)) return str;
  return convertPlainTextToHtml(str);
};

const buildTokenMap = (ctx) => {
  const js = ctx || {};
  const fmtDate = (val) => {
    if (!val) return '';
    const s = String(val);
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      try {
        const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
        return d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });
      } catch (_) {
        return s;
      }
    }
    return s;
  };
  const fmtCurrency = (val) => {
    if (val == null || val === '') return '';
    const num = Number(val);
    if (!Number.isFinite(num)) return '';
    try {
      return new Intl.NumberFormat(undefined, { style: 'currency', currency: 'GBP' }).format(num);
    } catch (_err) {
      return String(num);
    }
  };

  const firstName = (() => {
    const raw = String(js.client_name || '').trim();
    if (!raw) return '';
    const parts = raw.split(/\s+/);
    return parts[0] || '';
  })();

  return {
    client_name: js.client_name || '',
    client_first_name: firstName,
    client_email: js.client_email || '',
    event_type: js.event_type || '',
    event_date: fmtDate(js.event_date || ''),
    event_start: js.event_start || '',
    event_end: js.event_end || '',
    venue_name: js.venue_name || '',
    venue_town: js.venue_town || '',
    venue_postcode: js.venue_postcode || '',
    balance_amount: fmtCurrency(js.balance_amount),
    balance_due_date: fmtDate(js.balance_due_date || ''),
    today: fmtDate(new Date().toISOString().slice(0, 10))
  };
};

const createTokenChipHtml = (key, fallback, value) => {
  const safeValue = escapeHtml(value);
  const safeFallback = fallback ? escapeHtml(fallback) : '';
  return `<span class="${TOKEN_CHIP_CLASS}" data-token="${escapeHtml(key)}"${safeFallback ? ` data-fallback="${safeFallback}"` : ''} contenteditable="false" style="${TOKEN_CHIP_STYLE}" tabindex="-1">${safeValue || '&nbsp;'}</span>`;
};

const renderTemplateWithTokens = (templateHtml, tokenMap) => {
  const html = normalizeTemplateBody(templateHtml);
  const parser = new DOMParser();
  const doc = parser.parseFromString(html || '', 'text/html');
  const textNodes = [];
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null);
  let current;
  while ((current = walker.nextNode())) {
    textNodes.push(current);
  }
  textNodes.forEach(node => {
    const text = node.nodeValue || '';
    if (!TOKEN_REGEX.test(text)) return;
    TOKEN_REGEX.lastIndex = 0;
    const frag = doc.createDocumentFragment();
    let lastIndex = 0;
    text.replace(TOKEN_REGEX, (match, key, fallback, offset) => {
      if (offset > lastIndex) {
        frag.appendChild(doc.createTextNode(text.slice(lastIndex, offset)));
      }
      const normalizedKey = String(key || '').trim().toLowerCase();
      const fallbackValue = fallback != null ? String(fallback) : '';
      const resolved = tokenMap[normalizedKey] || '';
      const span = doc.createElement('span');
      span.className = TOKEN_CHIP_CLASS;
      span.setAttribute('contenteditable', 'false');
      span.setAttribute('tabindex', '-1');
      span.setAttribute('data-token', normalizedKey);
      if (fallbackValue) span.setAttribute('data-fallback', fallbackValue);
      span.setAttribute('style', TOKEN_CHIP_STYLE);
      span.textContent = resolved || fallbackValue || '';
      frag.appendChild(span);
      lastIndex = offset + match.length;
      return match;
    });
    if (lastIndex < text.length) {
      frag.appendChild(doc.createTextNode(text.slice(lastIndex)));
    }
    node.replaceWith(frag);
  });
  return doc.body.innerHTML;
};

const extractTemplateFromDisplay = (displayHtml) => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(displayHtml || '', 'text/html');
  doc.querySelectorAll(`span.${TOKEN_CHIP_CLASS}`).forEach(span => {
    const key = span.getAttribute('data-token');
    if (!key) return;
    const fallback = span.getAttribute('data-fallback') || '';
    const tokenString = `{{ ${key}${fallback ? `|${fallback}` : ''} }}`;
    const textNode = doc.createTextNode(tokenString);
    span.replaceWith(textNode);
  });
  return doc.body.innerHTML;
};

const appendSignatureHtml = (bodyHtml, signatureHtml) => {
  const trimmedBody = (bodyHtml || '').trim();
  if (!signatureHtml) return trimmedBody;
  if (!trimmedBody) return signatureHtml;
  if (/(<br\s*\/?>|<\/p>)$/i.test(trimmedBody)) {
    return `${trimmedBody}${signatureHtml}`;
  }
  return `${trimmedBody}<br><br>${signatureHtml}`;
};

const hasMeaningfulContent = (html) => {
  if (!html) return false;
  const normalized = normalizeTemplateBody(html);
  const stripped = normalized
    .replace(/<br\s*\/?>/gi, '')
    .replace(/<p>\s*<\/p>/gi, '')
    .replace(/<p[^>]*>|<\/p>/gi, '')
    .replace(/&nbsp;/gi, '')
    .trim();
  return Boolean(stripped);
};

const formatDateTimeLocal = (date) => {
  if (!(date instanceof Date) || Number.isNaN(date.valueOf())) return '';
  const pad = (value) => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
};

const computeDefaultScheduleDateTime = () => {
  const future = new Date(Date.now() + 15 * 60 * 1000);
  return formatDateTimeLocal(future);
};

export default function MailComposer({
  open,
  onClose,
  businessId,
  jobsheetId,
  initialTo = '',
  initialCc = '',
  initialBcc = '',
  initialSubject = '',
  initialBody = '',
  initialAttachments = [],
  initialTemplateKey = '',
  onTemplateChange,
  initialSendMode = 'now',
  initialScheduleAt = '',
  onSendModeChange,
  onScheduleChange,
  initialIncludeSignature = true,
  onIncludeSignatureChange,
  onSent
}) {
  const [to, setTo] = useState(initialTo);
  const [cc, setCc] = useState(initialCc);
  const [bcc, setBcc] = useState(initialBcc);
  const [templateSubject, setTemplateSubject] = useState(initialSubject || '');
  const [templateBody, setTemplateBody] = useState(() => normalizeTemplateBody(initialBody));
  const [attachments, setAttachments] = useState(initialAttachments || []);
  const [templates, setTemplates] = useState({});
  const [selectedTemplate, setSelectedTemplate] = useState(initialTemplateKey || '');
  const [signature, setSignature] = useState('');
  const [includeSignature, setIncludeSignature] = useState(initialIncludeSignature !== false);
  const [subjectDirty, setSubjectDirty] = useState(Boolean((initialSubject || '').trim()));
  const [bodyDirty, setBodyDirty] = useState(hasMeaningfulContent(initialBody));
  const [savingPreset, setSavingPreset] = useState(false);
  const [signatureModalOpen, setSignatureModalOpen] = useState(false);
  const [signatureDraft, setSignatureDraft] = useState('');
  const [signatureSaving, setSignatureSaving] = useState(false);
  const [dragging, setDragging] = useState(false);
  const dragStartRef = useRef({ x: 0, y: 0 });
  const [pos, setPos] = useState({ x: 0, y: 0 });
  const headerRef = useRef(null);
  const portalElRef = useRef(null);
  const subjectRef = useRef(null);
  const bodyEditorRef = useRef(null);
  const [lastFocus, setLastFocus] = useState('body');
  const [tokenChoice, setTokenChoice] = useState('client_name');
  const [jobFiles, setJobFiles] = useState([]);
  const [jobFilesLoading, setJobFilesLoading] = useState(false);
  const previousTemplateRef = useRef('');
  const [sendWhen, setSendWhen] = useState(() => (initialSendMode === 'later' ? 'later' : 'now'));
  const [scheduleDateTime, setScheduleDateTime] = useState(() => {
    if (initialSendMode === 'later' && initialScheduleAt) {
      const parsed = new Date(initialScheduleAt);
      if (!Number.isNaN(parsed.valueOf())) return formatDateTimeLocal(parsed);
    }
    return computeDefaultScheduleDateTime();
  });
  const [pendingBodyHtml, setPendingBodyHtml] = useState(null);
  const forceBodyReplaceRef = useRef(false);
  const firstApplyRef = useRef(true);
  const lastNotifiedSendModeRef = useRef(initialSendMode === 'later' ? 'later' : 'now');
  const lastNotifiedScheduleRef = useRef(initialScheduleAt || '');
  const lastNotifiedSignatureRef = useRef(initialIncludeSignature !== false);
  const selectedTemplateRef = useRef(initialTemplateKey || '');

  // Create a dedicated portal container and keep it last in <body>
  useEffect(() => {
    if (typeof document === 'undefined') return () => {};
    const el = document.createElement('div');
    // No positioning on the container so it never intercepts clicks when empty
    document.body.appendChild(el);
    portalElRef.current = el;
    return () => {
      if (el && el.parentNode) el.parentNode.removeChild(el);
      portalElRef.current = null;
    };
  }, []);

  const bringPortalToFront = () => {
    const el = portalElRef.current;
    if (el && el.parentNode) el.parentNode.appendChild(el);
  };

  useEffect(() => {
    if (open) bringPortalToFront();
  }, [open, businessId, jobsheetId, initialTemplateKey]);

  useEffect(() => {
    if (!dragging) return;
    const onMove = (e) => {
      e.preventDefault();
      const ds = dragStartRef.current || { x: 0, y: 0 };
      setPos({ x: e.clientX - ds.x, y: e.clientY - ds.y });
    };
    const onUp = () => {
      setDragging(false);
      window.removeEventListener('mousemove', onMove, true);
      window.removeEventListener('mouseup', onUp, true);
      window.removeEventListener('blur', onUp, true);
    };
    window.addEventListener('mousemove', onMove, true);
    window.addEventListener('mouseup', onUp, true);
    window.addEventListener('blur', onUp, true);
    return () => {
      window.removeEventListener('mousemove', onMove, true);
      window.removeEventListener('mouseup', onUp, true);
      window.removeEventListener('blur', onUp, true);
    };
  }, [dragging]);

  // Persist position between sessions (localStorage)
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem('invoiceMaster:mailComposerPos');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed.x === 'number' && typeof parsed.y === 'number') {
          setPos({ x: parsed.x, y: parsed.y });
        }
      }
    } catch (_) {}
  }, []);

  useEffect(() => {
    if (!dragging) return;
    const save = () => {
      try { window.localStorage.setItem('invoiceMaster:mailComposerPos', JSON.stringify(pos)); } catch (_) {}
    };
    const onUp = () => save();
    window.addEventListener('mouseup', onUp, true);
    window.addEventListener('blur', onUp, true);
    return () => {
      window.removeEventListener('mouseup', onUp, true);
      window.removeEventListener('blur', onUp, true);
    };
  }, [dragging, pos]);
  const [busy, setBusy] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [templateCtx, setTemplateCtx] = useState({});

  useEffect(() => {
    selectedTemplateRef.current = selectedTemplate;
  }, [selectedTemplate]);

  useEffect(() => {
    if (!open) return;
    setTo(initialTo);
    setCc(initialCc);
    setBcc(initialBcc);
    setTemplateSubject(initialSubject || '');
    setTemplateBody(normalizeTemplateBody(initialBody));
    setAttachments(Array.isArray(initialAttachments) ? [...initialAttachments] : []);
    setSelectedTemplate(initialTemplateKey || '');
    const normalizedMode = initialSendMode === 'later' ? 'later' : 'now';
    setSendWhen(prev => (prev === normalizedMode ? prev : normalizedMode));
    if (normalizedMode === 'later') {
      if (initialScheduleAt) {
        const parsed = new Date(initialScheduleAt);
        if (!Number.isNaN(parsed.valueOf())) {
          const next = formatDateTimeLocal(parsed);
          setScheduleDateTime(next);
          lastNotifiedScheduleRef.current = parsed.toISOString();
        } else {
          const fallback = computeDefaultScheduleDateTime();
          setScheduleDateTime(fallback);
          lastNotifiedScheduleRef.current = '';
        }
      } else {
        const fallback = computeDefaultScheduleDateTime();
        setScheduleDateTime(fallback);
        lastNotifiedScheduleRef.current = '';
      }
    } else {
      lastNotifiedScheduleRef.current = '';
    }
    const normalizedInclude = initialIncludeSignature !== false;
    setIncludeSignature(prev => (prev === normalizedInclude ? prev : normalizedInclude));
    lastNotifiedSendModeRef.current = normalizedMode;
    lastNotifiedSignatureRef.current = normalizedInclude;
    selectedTemplateRef.current = initialTemplateKey || '';
    setSubjectDirty(Boolean((initialSubject || '').trim()));
    setBodyDirty(hasMeaningfulContent(initialBody));
    // When parent provides a new initial body (e.g., switching template programmatically),
    // force the contentEditable to refresh even if focused.
    forceBodyReplaceRef.current = true;
  }, [open, initialTo, initialCc, initialBcc, initialSubject, initialBody, initialAttachments, initialTemplateKey, initialSendMode, initialScheduleAt, initialIncludeSignature]);

  // always HTML; no toggle needed

  const pushToast = (text, tone = 'info') => {
    const notice = { id: `toast-${Date.now()}-${Math.random().toString(36).slice(2)}`, text, tone };
    setToasts(prev => [...prev, notice]);
    setTimeout(() => setToasts(prev => prev.filter(t => t !== notice)), 3500);
  };

  useEffect(() => {
    if (!open) return;
    let mounted = true;
    (async () => {
      let fetchedTemplates = {};
      try {
        const [tplResult, defaultResult, signatureResult, tombstones] = await Promise.all([
          window.api?.getMailTemplates?.({ businessId }),
          window.api?.getDefaultMailTemplates?.({ businessId }),
          window.api?.getMailSignature?.({ businessId }),
          window.api?.getMailTemplateTombstones?.({ businessId })
        ]);
        if (!mounted) return;
        const tomb = Array.isArray(tombstones) ? new Set(tombstones.map(k => String(k || '').toLowerCase())) : new Set();
        const defs = defaultResult || {};
        const custom = tplResult || {};
        const keys = new Set([...Object.keys(defs), ...Object.keys(custom)]);
        const nonEmpty = (v) => v != null && String(v).trim() !== '';
        const mergedMap = {};
        keys.forEach(k => {
          const kl = String(k).toLowerCase();
          if (tomb.has(kl)) return; // respect deletions
          const d = defs[k] || {};
          const c = custom[k] || {};
          mergedMap[k] = {
            label: nonEmpty(c.label) ? c.label : (d.label || k),
            subject: nonEmpty(c.subject) ? c.subject : (d.subject || ''),
            body: nonEmpty(c.body) ? c.body : (d.body || '')
          };
        });
        fetchedTemplates = mergedMap;
        setTemplates(fetchedTemplates);
        const sig = (signatureResult && signatureResult.signature) || '';
        setSignature(sig);
        setSignatureDraft(sig);
      } catch (_) {}
      // Load template context from jobsheet if available
      try {
        if (jobsheetId && window.api?.getAhmenJobsheet) {
          const js = await window.api.getAhmenJobsheet(jobsheetId);
          if (mounted && js) setTemplateCtx(js || {});
        } else {
          if (mounted) setTemplateCtx({});
        }
      } catch (_) { if (mounted) setTemplateCtx({}); }
      // Load job folder files
      try {
        setJobFilesLoading(true);
        const files = await window.api?.listJobFolderFiles?.({ businessId, jobsheetId, extensionPattern: '\\.(pdf)$' });
        if (mounted) setJobFiles(Array.isArray(files) ? files : []);
      } catch (_) { if (mounted) setJobFiles([]); } finally { if (mounted) setJobFilesLoading(false); }
      if (!mounted) return;
      const keys = Object.keys(fetchedTemplates || {});
      if (!keys.length) return;
      let nextKey = '';
      if (initialTemplateKey && fetchedTemplates[initialTemplateKey]) {
        nextKey = initialTemplateKey;
      } else if (selectedTemplateRef.current && fetchedTemplates[selectedTemplateRef.current]) {
        nextKey = selectedTemplateRef.current;
      } else {
        nextKey = keys.includes('enquiry_ack') ? 'enquiry_ack' : keys[0];
      }
      if (nextKey) {
        const tpl = fetchedTemplates[nextKey] || {};
        if (!initialSubject || !String(initialSubject).trim()) {
          const nextSubject = tpl.subject || '';
          setTemplateSubject(prev => (prev === nextSubject ? prev : nextSubject));
        }
        if (!hasMeaningfulContent(initialBody)) {
          const nextBody = normalizeTemplateBody(tpl.body);
          setTemplateBody(prev => (prev === nextBody ? prev : nextBody));
          forceBodyReplaceRef.current = true;
        }
        if (nextKey !== selectedTemplateRef.current) {
          selectedTemplateRef.current = nextKey;
          setSelectedTemplate(nextKey);
        }
      }
    })();
    return () => { mounted = false; };
  }, [open, businessId, jobsheetId, initialTemplateKey]);

  useEffect(() => {
    if (!open) return;
    const tpl = templates[selectedTemplate];
    if (!tpl) return;
    // On first apply for this mount, aggressively apply template values if caller didn't provide content
    if (firstApplyRef.current) {
      let changed = false;
      if (!initialSubject || !String(initialSubject).trim()) {
        const nextSubject = tpl.subject || '';
        setTemplateSubject(prev => (prev === nextSubject ? prev : nextSubject));
        changed = true;
      }
      if (!hasMeaningfulContent(initialBody)) {
        const nextBody = normalizeTemplateBody(tpl.body);
        setTemplateBody(prev => (prev === nextBody ? prev : nextBody));
        forceBodyReplaceRef.current = true;
        changed = true;
      }
      if (changed) {
        firstApplyRef.current = false;
        return;
      }
      firstApplyRef.current = false;
    }
    if (!subjectDirty) {
      const nextSubject = tpl.subject || '';
      setTemplateSubject(prev => (prev === nextSubject ? prev : nextSubject));
    }
    if (!bodyDirty) {
      const nextBody = normalizeTemplateBody(tpl.body);
      setTemplateBody(prev => (prev === nextBody ? prev : nextBody));
    }
  }, [open, templates, selectedTemplate, subjectDirty, bodyDirty, initialSubject, initialBody]);

  useEffect(() => {
    if (typeof onTemplateChange !== 'function') return;
    onTemplateChange(selectedTemplate || '');
  }, [selectedTemplate, onTemplateChange]);

  useEffect(() => {
    if (typeof onSendModeChange === 'function' && sendWhen !== lastNotifiedSendModeRef.current) {
      lastNotifiedSendModeRef.current = sendWhen;
      onSendModeChange(sendWhen);
    }
  }, [sendWhen, onSendModeChange]);

  useEffect(() => {
    if (typeof onScheduleChange !== 'function') return;
    if (!scheduleDateTime) {
      onScheduleChange('');
      return;
    }
    const parsed = new Date(scheduleDateTime);
    if (Number.isNaN(parsed.valueOf())) {
      if (lastNotifiedScheduleRef.current !== '') {
        lastNotifiedScheduleRef.current = '';
        onScheduleChange('');
      }
      return;
    }
    const iso = parsed.toISOString();
    if (lastNotifiedScheduleRef.current !== iso) {
      lastNotifiedScheduleRef.current = iso;
      onScheduleChange(iso);
    }
  }, [scheduleDateTime, onScheduleChange]);

  useEffect(() => {
    if (typeof onIncludeSignatureChange === 'function' && includeSignature !== lastNotifiedSignatureRef.current) {
      lastNotifiedSignatureRef.current = includeSignature;
      onIncludeSignatureChange(includeSignature);
    }
  }, [includeSignature, onIncludeSignatureChange]);

  const handleTemplateSelect = useCallback((key, sourceTemplates = templates) => {
    if (!key || !sourceTemplates?.[key]) return;
    setSubjectDirty(false);
    setBodyDirty(false);
    setSelectedTemplate(key);
  }, [templates]);

  const tokenMap = useMemo(() => buildTokenMap(templateCtx), [templateCtx]);

  const applyTemplate = useCallback((text) => {
    if (!text) return '';
    return String(text).replace(TOKEN_REGEX, (_m, key, defVal) => {
      const normalizedKey = String(key || '').trim().toLowerCase();
      const value = tokenMap[normalizedKey];
      if (value != null && value !== '') return String(value);
      return defVal != null ? String(defVal) : '';
    });
  }, [tokenMap]);

  const displayBodyHtml = useMemo(() => renderTemplateWithTokens(templateBody, tokenMap), [templateBody, tokenMap]);

  useEffect(() => {
    if (!open) return;
    const el = bodyEditorRef.current;
    if (!el) return;
    const html = displayBodyHtml || '<p><br></p>';
    const force = forceBodyReplaceRef.current === true;
    if (document.activeElement === el && !force) {
      setPendingBodyHtml(html);
      return;
    }
    if (el.innerHTML !== html) {
      el.innerHTML = html;
    }
    if (force) {
      forceBodyReplaceRef.current = false;
      setPendingBodyHtml(null);
    }
  }, [open, displayBodyHtml]);

  useEffect(() => {
    if (!pendingBodyHtml) return;
    const el = bodyEditorRef.current;
    if (!el) return;
    if (document.activeElement === el) return;
    if (el.innerHTML !== pendingBodyHtml) {
      el.innerHTML = pendingBodyHtml;
    }
    setPendingBodyHtml(null);
  }, [pendingBodyHtml]);

  useEffect(() => {
    const el = bodyEditorRef.current;
    if (!el) return () => {};
    const onBlur = () => {
      if (!pendingBodyHtml) return;
      if (el.innerHTML !== pendingBodyHtml) {
        el.innerHTML = pendingBodyHtml;
      }
      setPendingBodyHtml(null);
    };
    el.addEventListener('blur', onBlur);
    return () => el.removeEventListener('blur', onBlur);
  }, [pendingBodyHtml]);

  useEffect(() => {
    if (!open) {
      previousTemplateRef.current = '';
    }
  }, [open]);

  useEffect(() => {
    if (!open) return;
    if (!selectedTemplate) return;
    if (!businessId) return;
    const normalizedKey = String(selectedTemplate).toLowerCase();
    if (previousTemplateRef.current === normalizedKey) return;
    previousTemplateRef.current = normalizedKey;

    let cancelled = false;
    (async () => {
      try {
        const res = await window.api?.resolveTemplateDefaultAttachments?.({
          businessId,
          jobsheetId,
          templateKey: normalizedKey
        });
        if (cancelled) return;
        const defaults = Array.isArray(res?.attachments) ? res.attachments.filter(Boolean) : [];
        if (defaults.length === 0) {
          setAttachments([]);
        } else {
          setAttachments(defaults.map(p => String(p)));
        }
      } catch (_) {}
    })();

    return () => {
      cancelled = true;
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, selectedTemplate, businessId, jobsheetId]);

  const handleBodyInput = useCallback(() => {
    const el = bodyEditorRef.current;
    if (!el) return;
    const html = el.innerHTML;
    const templateHtml = normalizeTemplateBody(extractTemplateFromDisplay(html));
    setTemplateBody(templateHtml);
    setBodyDirty(true);
  }, []);

  const resolveTokenValue = useCallback((key, fallback) => {
    const normalizedKey = String(key || '').trim().toLowerCase();
    const value = tokenMap[normalizedKey];
    if (value != null && value !== '') return String(value);
    return fallback || '';
  }, [tokenMap]);

  const insertTokenIntoSubject = useCallback((tokenKey, fallback) => {
    const tokenString = `{{ ${tokenKey}${fallback ? `|${fallback}` : ''} }}`;
    const input = subjectRef.current;
    if (input && typeof input.selectionStart === 'number') {
      const start = input.selectionStart;
      const end = input.selectionEnd ?? start;
      const current = String(templateSubject || '');
      const next = current.slice(0, start) + tokenString + current.slice(end);
      setTemplateSubject(next);
      setSubjectDirty(true);
      requestAnimationFrame(() => {
        try {
          input.focus();
          const pos = start + tokenString.length;
          input.setSelectionRange(pos, pos);
        } catch (_) {}
      });
    } else {
      setTemplateSubject(prev => {
        setSubjectDirty(true);
        return `${prev || ''}${tokenString}`;
      });
    }
  }, [templateSubject]);

  const insertTokenIntoBody = useCallback((tokenKey, fallback) => {
    const el = bodyEditorRef.current;
    if (!el) return;
    el.focus();
    const resolved = resolveTokenValue(tokenKey, fallback);
    const chipHtml = createTokenChipHtml(tokenKey, fallback, resolved);
    document.execCommand('insertHTML', false, chipHtml);
    handleBodyInput();
  }, [resolveTokenValue, handleBodyInput]);

  const insertToken = useCallback((tokenKey) => {
    const fallback = TOKEN_FALLBACKS[tokenKey] || '';
    if (lastFocus === 'subject') {
      insertTokenIntoSubject(tokenKey, fallback);
    } else {
      insertTokenIntoBody(tokenKey, fallback);
    }
  }, [lastFocus, insertTokenIntoSubject, insertTokenIntoBody]);

  const bodyWithSignatureTokens = useMemo(() => appendSignatureHtml(templateBody, includeSignature ? signature : ''), [templateBody, includeSignature, signature]);
  const renderedBodyFinal = useMemo(() => applyTemplate(bodyWithSignatureTokens), [bodyWithSignatureTokens, applyTemplate]);
  const renderedSubject = useMemo(() => applyTemplate(templateSubject), [templateSubject, applyTemplate]);

  const openSignatureModal = useCallback(() => {
    bringPortalToFront();
    setSignatureDraft(signature || '');
    setSignatureModalOpen(true);
  }, [signature]);

  const closeSignatureModal = useCallback(() => {
    setSignatureModalOpen(false);
  }, []);

  const handleSaveSignature = useCallback(async () => {
    try {
      setSignatureSaving(true);
      await window.api?.saveMailSignature?.({ businessId, signature: signatureDraft });
      setSignature(signatureDraft);
      pushToast('Signature updated', 'success');
      setSignatureModalOpen(false);
    } catch (err) {
      pushToast(err?.message || 'Unable to save signature', 'error');
    } finally {
      setSignatureSaving(false);
    }
  }, [businessId, signatureDraft]);

  const handleSend = async () => {
    if (!to.trim()) { window.alert('Enter recipient'); return; }
    const finalSubject = renderedSubject;
    const finalBody = renderedBodyFinal;
    let scheduledAt = null;
    if (sendWhen === 'later') {
      if (!scheduleDateTime) { window.alert('Select a schedule date and time'); return; }
      scheduledAt = new Date(scheduleDateTime);
      if (Number.isNaN(scheduledAt.valueOf())) { window.alert('Schedule time is invalid'); return; }
      if (scheduledAt.getTime() < Date.now() + 30 * 1000) {
        window.alert('Scheduled time must be at least 30 seconds in the future');
        return;
      }
    }

    try {
      setBusy(true);
      if (sendWhen === 'later') {
        await window.api?.scheduleMailViaGraph?.({
          to,
          cc,
          bcc,
          subject: finalSubject,
          body: finalBody,
          attachments,
          is_html: true,
          business_id: businessId,
          jobsheet_id: jobsheetId,
          send_at: scheduledAt.toISOString()
        });
      } else {
        await window.api?.sendMailViaGraph?.({
          to,
          cc,
          bcc,
          subject: finalSubject,
          body: finalBody,
          attachments,
          is_html: true,
          business_id: businessId,
          jobsheet_id: jobsheetId
        });
      }
      try {
        window.api?.notifyJobsheetChange?.({
          type: 'email-log-updated',
          businessId,
          jobsheetId
        });
      } catch (_) {}
      onClose?.();
      onSent?.({ mode: sendWhen });
    } catch (err) {
      window.alert(err?.message || (sendWhen === 'later' ? 'Unable to schedule email' : 'Unable to send email'));
    } finally {
      setBusy(false);
    }
  };

  if (!open) return null;

  const composerContent = (
    <div
      className="fixed inset-0 bg-slate-900/50 p-4"
      style={{
        zIndex: signatureModalOpen ? 9999999998 : 9999999999,
        pointerEvents: signatureModalOpen ? 'none' : 'auto'
      }}
    >
      <div
        className="w-full max-w-3xl rounded-lg bg-white shadow-2xl ring-2 ring-indigo-500/40"
        style={{
          position: 'fixed',
          left: Math.max(8, pos.x),
          top: Math.max(8, pos.y),
          zIndex: signatureModalOpen ? 9999999998 : 10000000000,
          pointerEvents: signatureModalOpen ? 'none' : 'auto'
        }}
      >
        <div
          ref={headerRef}
          className="flex items-center justify-between border-b border-indigo-600/20 px-4 py-3 cursor-move select-none bg-indigo-600 text-white rounded-t-lg"
          onMouseDown={(e) => {
            bringPortalToFront();
            const targetNode = e.target;
            const targetEl = targetNode instanceof Element
              ? targetNode
              : (targetNode && targetNode.parentElement ? targetNode.parentElement : null);
            if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('button')) {
              return;
            }
            e.preventDefault();
            dragStartRef.current = { x: e.clientX - pos.x, y: e.clientY - pos.y };
            setDragging(true);
          }}
        >
          <h3 className="text-base font-semibold">Compose email</h3>
          <button
            className="text-white/80 hover:text-white"
            onClick={onClose}
            onMouseDown={event => event.stopPropagation()}
            aria-label="Close"
          >✕</button>
        </div>
        <div className="p-4 space-y-3 text-sm">
          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">Template</label>
            <div className="col-span-5 flex items-center gap-2">
              <select value={selectedTemplate} onChange={e => handleTemplateSelect(e.target.value)} className="rounded border border-slate-300 px-2 py-1">
                {Object.keys(templates || {}).length === 0 ? (
                  <option value="" disabled>(no templates)</option>
                ) : (
                  Object.keys(templates || {}).map(k => (
                    <option key={k} value={k}>{templates[k]?.label || k}</option>
                  ))
                )}
              </select>
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1 text-xs"
                disabled={savingPreset}
                onClick={async () => {
                  try {
                    setSavingPreset(true);
                    const next = { ...(templates || {}) };
                    let key = selectedTemplate;
                    if (!key) {
                      const name = window.prompt('Template name', 'New template');
                      if (!name) { setSavingPreset(false); return; }
                      key = name.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
                      if (!key) { setSavingPreset(false); return; }
                    }
                    const label = next[key]?.label || key;
                    next[key] = { subject: templateSubject, body: templateBody, label };
                    await window.api?.saveMailTemplates?.({ businessId, templates: next });
                    setTemplates(next);
                    handleTemplateSelect(key, next);
                    pushToast('Template saved', 'success');
                  } catch (err) {
                    pushToast(err?.message || 'Unable to save template', 'error');
                  } finally {
                    setSavingPreset(false);
                  }
                }}
              >Save template</button>
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1 text-xs"
                onClick={async () => {
                  try {
                    const defaults = await window.api?.getDefaultMailTemplates?.({ businessId });
                    const tomb = await window.api?.getMailTemplateTombstones?.({ businessId });
                    const existing = templates || {};
                    const hasAny = Object.keys(existing).length > 0;
                    let replace = false;
                    if (hasAny) {
                      replace = window.confirm('Replace existing templates with defaults? Click OK to replace, Cancel to add missing only.');
                    }
                    let next = {};
                    if (replace) {
                      next = { ...existing, ...defaults };
                    } else {
                      next = { ...defaults, ...existing };
                    }
                    // Respect tombstones: never add deleted default keys
                    if (Array.isArray(tomb) && tomb.length) {
                      const hide = new Set(tomb.map(k => String(k || '').toLowerCase()));
                      Object.keys(next).forEach(k => { if (hide.has(String(k).toLowerCase())) delete next[k]; });
                    }
                    await window.api?.saveMailTemplates?.({ businessId, templates: next });
                    setTemplates(next);
                    const pref = selectedTemplate && next[selectedTemplate]
                      ? selectedTemplate
                      : (Object.keys(next).includes('enquiry_ack') ? 'enquiry_ack' : Object.keys(next)[0]);
                    if (pref) handleTemplateSelect(pref, next);
                    pushToast('Defaults seeded', 'success');
                  } catch (err) {
                    pushToast(err?.message || 'Unable to seed defaults', 'error');
                  }
                }}
              >Seed defaults</button>
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1 text-xs"
                onClick={async () => {
                  const name = window.prompt('New template name (e.g. “Chaser”)', '');
                  if (!name) return;
                  const key = name.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
                  if (!key) return;
                  const next = { ...(templates || {}) };
                  if (next[key]) { pushToast('Template already exists', 'warning'); return; }
                  next[key] = { label: name, subject: templateSubject, body: templateBody };
                  try { await window.api?.saveMailTemplates?.({ businessId, templates: next }); setTemplates(next); handleTemplateSelect(key, next); pushToast('Template created', 'success'); } catch (err) { pushToast(err?.message || 'Unable to create template', 'error'); }
                }}
              >New template…</button>
              {selectedTemplate && templates[selectedTemplate] ? (
                <button
                  type="button"
                  className="rounded border border-red-300 text-red-700 px-2 py-1 text-xs"
                  onClick={async () => {
                    if (!window.confirm('Delete this template?')) return;
                    try {
                      await window.api?.deleteMailTemplate?.({ businessId, key: selectedTemplate });
                      // Remove locally and advance selection
                      const next = { ...(templates || {}) };
                      delete next[selectedTemplate];
                      setTemplates(next);
                      const first = Object.keys(next).includes('enquiry_ack') ? 'enquiry_ack' : (Object.keys(next)[0] || '');
                      if (first) handleTemplateSelect(first, next);
                      else setSelectedTemplate('');
                      pushToast('Template deleted', 'success');
                    } catch (err) { pushToast(err?.message || 'Unable to delete template', 'error'); }
                  }}
                >Delete</button>
              ) : null}
            </div>
          </div>

          {/* Job folder files with inline checkboxes */}
          <div className="grid grid-cols-6 gap-2 items-start">
            <label className="col-span-1 text-slate-600 mt-1">Job files</label>
            <div className="col-span-5">
              <div className="mb-2 flex items-center gap-2 text-xs">
                <button type="button" className="rounded border border-slate-300 px-2 py-1" onClick={async () => {
                  try { setJobFilesLoading(true); const files = await window.api?.listJobFolderFiles?.({ businessId, jobsheetId, extensionPattern: '\\.(pdf)$' }); setJobFiles(Array.isArray(files) ? files : []); } catch (_) {} finally { setJobFilesLoading(false); }
                }}>Refresh</button>
                <span className="text-slate-500">Tick to attach from the job folder</span>
              </div>
              <div className="max-h-48 overflow-auto rounded border border-slate-200">
                {jobFilesLoading ? (
                  <div className="px-2 py-2 text-sm text-slate-500">Loading…</div>
                ) : (jobFiles.length === 0 ? (
                  <div className="px-2 py-2 text-sm text-slate-500">No files in job folder</div>
                ) : jobFiles.map(f => {
                  const checked = attachments.includes(f.path);
                  return (
                    <label key={f.path} className="flex items-center gap-2 px-2 py-1 text-sm text-slate-700">
                      <input type="checkbox" checked={checked} onChange={e => {
                        const on = e.target.checked;
                        setAttachments(prev => {
                          const set = new Set(prev || []);
                          if (on) set.add(f.path); else set.delete(f.path);
                          return Array.from(set);
                        });
                      }} />
                      <span className="truncate" title={f.path}>{f.name}</span>
                      <span className="ml-auto text-xs text-slate-500">{Math.round((f.size || 0) / 1024)} KB</span>
                    </label>
                  );
                }))}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">Tokens</label>
            <div className="col-span-5 flex items-center gap-2">
              <select value={tokenChoice} onChange={e => setTokenChoice(e.target.value)} className="rounded border border-slate-300 px-2 py-1">
                {TOKEN_OPTIONS.map(opt => (
                  <option key={opt.key} value={opt.key}>{opt.label}</option>
                ))}
              </select>
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1 text-xs"
                title="Focus Subject or Body, then insert"
                onClick={() => insertToken(tokenChoice)}
              >Insert token</button>
              <span className="text-xs text-slate-500">Inserts at cursor in last focused field</span>
            </div>
          </div>

          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">To</label>
            <input className="col-span-5 rounded border border-slate-300 px-2 py-1" value={to} onChange={e => setTo(e.target.value)} placeholder="email@example.com" />
          </div>
          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">Cc</label>
            <input className="col-span-5 rounded border border-slate-300 px-2 py-1" value={cc} onChange={e => setCc(e.target.value)} placeholder="optional" />
          </div>
          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">Bcc</label>
            <input className="col-span-5 rounded border border-slate-300 px-2 py-1" value={bcc} onChange={e => setBcc(e.target.value)} placeholder="optional" />
          </div>
          <div className="grid grid-cols-6 gap-2 items-center">
            <label className="col-span-1 text-slate-600">Subject</label>
            <div className="col-span-5">
              <input
                ref={subjectRef}
                className="w-full rounded border border-slate-300 px-2 py-1"
                value={templateSubject}
                onChange={e => {
                  setTemplateSubject(e.target.value);
                  setSubjectDirty(true);
                }}
                onFocus={() => setLastFocus('subject')}
              />
              <div className="mt-1 text-xs text-slate-500">Preview: {renderedSubject || '(empty)'}</div>
            </div>
          </div>
          <div>
            <label className="block text-slate-600 mb-1">Body</label>
            <div
              ref={bodyEditorRef}
              className="rounded border border-slate-300 px-2 py-2 whitespace-pre-wrap focus:outline-none focus:ring-2 focus:ring-indigo-500"
              style={{ minHeight: '160px' }}
              contentEditable
              suppressContentEditableWarning
              onInput={handleBodyInput}
              onFocus={() => setLastFocus('body')}
            />
            <div className="mt-2 text-xs text-slate-500">Tokens show live data. Delete a chip to remove a token.</div>
            <div className="mt-2 flex items-center gap-2 text-xs">
              <label className="inline-flex items-center gap-2 text-slate-600">
                <input type="checkbox" checked={includeSignature} onChange={e => setIncludeSignature(e.target.checked)} />
                Include signature
              </label>
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1"
                onClick={openSignatureModal}
              >Edit signature…</button>
            </div>
          </div>
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <div className="text-slate-600">Attachments</div>
              
              <button
                type="button"
                className="rounded border border-slate-300 px-2 py-1 text-xs"
                onClick={() => setAttachments([])}
              >Clear all</button>
            </div>
            <div className="rounded border border-slate-200 divide-y">
              {attachments.length === 0 ? (
                <div className="px-2 py-2 text-slate-500 text-sm">No attachments</div>
              ) : attachments.map((p, idx) => (
                <div key={`${p}-${idx}`} className="flex items-center justify-between px-2 py-1">
                  <div className="truncate text-sm text-slate-700" title={p}>{String(p).split(/[\\/]+/).pop()}</div>
                  <button type="button" className="rounded border border-slate-300 px-2 py-0.5 text-xs" onClick={() => setAttachments(prev => prev.filter(x => x !== p))}>Remove</button>
                </div>
              ))}
            </div>
          </div>
          <div className="space-y-2">
            <div className="text-slate-600">Send timing</div>
            <div className="flex items-center gap-4 text-sm text-slate-600">
              <label className="inline-flex items-center gap-2">
                <input
                  type="radio"
                  name="mail-send-when"
                  value="now"
                  checked={sendWhen === 'now'}
                  onChange={() => setSendWhen('now')}
                  disabled={busy}
                />
                Send now
              </label>
              <label className="inline-flex items-center gap-2">
                <input
                  type="radio"
                  name="mail-send-when"
                  value="later"
                  checked={sendWhen === 'later'}
                  onChange={() => {
                    setSendWhen('later');
                    const current = scheduleDateTime ? new Date(scheduleDateTime) : null;
                    if (!scheduleDateTime || Number.isNaN(current?.valueOf()) || current.getTime() < Date.now()) {
                      setScheduleDateTime(computeDefaultScheduleDateTime());
                    }
                  }}
                  disabled={busy}
                />
                Schedule send
              </label>
            </div>
            {sendWhen === 'later' ? (
              <div className="flex flex-col gap-1">
                <input
                  type="datetime-local"
                  className="w-full rounded border border-slate-300 px-2 py-1 text-sm"
                  value={scheduleDateTime}
                  onChange={e => setScheduleDateTime(e.target.value)}
                  min={formatDateTimeLocal(new Date(Date.now() + 60 * 1000))}
                  disabled={busy}
                />
                <span className="text-xs text-slate-500">Times use your local time zone.</span>
              </div>
            ) : null}
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 border-t border-slate-200 px-4 py-3">
          <button className="rounded border border-slate-300 px-3 py-1.5 text-sm" onClick={onClose}>Cancel</button>
          <button
            className="rounded bg-green-600 px-3 py-1.5 text-sm font-semibold text-white hover:bg-green-500 disabled:opacity-60"
            disabled={busy}
            onClick={handleSend}
          >{sendWhen === 'later' ? 'Schedule' : 'Send'}</button>
        </div>
      </div>
      <ToastOverlay notices={toasts} />
    </div>
  );

  const signatureContent = signatureModalOpen ? (
    <div className="fixed inset-0 bg-slate-900/60 px-4 py-6 flex items-center justify-center" style={{ zIndex: 20000000000 }}>
      <div className="w-full max-w-2xl rounded-lg bg-white shadow-2xl">
        <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
          <h3 className="text-base font-semibold text-slate-800">Edit signature</h3>
          <button
            className="text-slate-400 hover:text-slate-600"
            onClick={closeSignatureModal}
            onMouseDown={event => event.stopPropagation()}
            aria-label="Close"
          >✕</button>
        </div>
        <div className="p-4 space-y-3 text-sm">
          <textarea
            rows={10}
            className="w-full rounded border border-slate-300 px-2 py-1 font-mono text-xs"
            value={signatureDraft}
            onChange={e => setSignatureDraft(e.target.value)}
            placeholder="Paste or edit your HTML signature here"
          />
          <p className="text-xs text-slate-500">
            HTML is supported. Use hosted image URLs (e.g. Dropbox links) inside &lt;img&gt; tags to display logos.
          </p>
        </div>
        <div className="flex items-center justify-end gap-2 border-t border-slate-200 px-4 py-3">
          <button className="rounded border border-slate-300 px-3 py-1.5 text-sm" onClick={closeSignatureModal}>Cancel</button>
          <button
            className="rounded bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white hover:bg-indigo-500 disabled:opacity-60"
            onClick={handleSaveSignature}
            disabled={signatureSaving}
          >{signatureSaving ? 'Saving…' : 'Save'}</button>
        </div>
      </div>
    </div>
  ) : null;

  if (typeof document === 'undefined') {
    return (
      <>
        {composerContent}
        {signatureContent}
      </>
    );
  }

  const target = portalElRef.current || document.body;
  const composerPortal = createPortal(composerContent, target);
  const signaturePortal = signatureModalOpen ? createPortal(signatureContent, document.body) : null;
  return (
    <>
      {composerPortal}
      {signaturePortal}
    </>
  );
}

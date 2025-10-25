import React, { useEffect, useRef } from 'react';

function WysiwygEditor({ value, onChange, height = 360 }) {
  const ref = useRef(null);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const current = el.innerHTML;
    if (value != null && String(value) !== current) {
      el.innerHTML = value || '';
    }
  }, [value]);

  const exec = (cmd, arg = null) => {
    try { document.execCommand('styleWithCSS', false, true); } catch (_) {}
    document.execCommand(cmd, false, arg);
    if (typeof onChange === 'function') {
      onChange(ref.current ? ref.current.innerHTML : '');
    }
  };

  const onInput = () => {
    if (typeof onChange === 'function') onChange(ref.current ? ref.current.innerHTML : '');
  };

  const onPaste = (e) => {
    // Clean paste: strip formatting except basic
    try {
      e.preventDefault();
      const text = e.clipboardData.getData('text/plain');
      document.execCommand('insertText', false, text);
    } catch (_) {}
  };

  return (
    <div className="border rounded">
      <div className="flex flex-wrap gap-1 border-b px-2 py-1 bg-slate-50">
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('bold')}>B</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('italic')}>I</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('underline')}>U</button>
        <span className="mx-1" />
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('formatBlock', 'H1')}>H1</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('formatBlock', 'H2')}>H2</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('formatBlock', 'P')}>P</button>
        <span className="mx-1" />
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('insertUnorderedList')}>• List</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('insertOrderedList')}>1. List</button>
        <span className="mx-1" />
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => { const url = prompt('Link URL'); if (url) exec('createLink', url); }}>Link</button>
        <span className="mx-1" />
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('justifyLeft')}>Left</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('justifyCenter')}>Center</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('justifyRight')}>Right</button>
        <span className="mx-1" />
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('undo')}>Undo</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('redo')}>Redo</button>
        <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => exec('removeFormat')}>Clear</button>
      </div>
      <div
        ref={ref}
        onInput={onInput}
        onPaste={onPaste}
        contentEditable
        suppressContentEditableWarning
        className="px-3 py-2 text-sm"
        style={{ minHeight: height, outline: 'none' }}
      />
    </div>
  );
}

export default WysiwygEditor;


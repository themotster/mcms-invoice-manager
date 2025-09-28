import React from 'react';
import { createPortal } from 'react-dom';

const TONE_CLASSES = {
  error: 'border-red-200 bg-red-50 text-red-700',
  success: 'border-green-200 bg-green-50 text-green-700',
  warning: 'border-amber-200 bg-amber-50 text-amber-700',
  info: 'border-indigo-200 bg-indigo-50 text-indigo-700'
};

function resolveTone(tone) {
  if (!tone) return 'info';
  const normalized = tone.toLowerCase();
  if (TONE_CLASSES[normalized]) return normalized;
  return 'info';
}

export default function ToastOverlay({ notices }) {
  const items = Array.isArray(notices)
    ? notices.filter(item => item && typeof item.text === 'string' && item.text.trim())
    : [];
  if (!items.length) return null;

  const content = (
    <div className="pointer-events-none fixed inset-x-0 bottom-4 z-[2000] flex justify-center px-4">
      <div className="flex w-full max-w-3xl flex-col gap-2">
        {items.map((item, index) => {
          const tone = resolveTone(item.tone);
          const className = TONE_CLASSES[tone] || TONE_CLASSES.info;
          const key = item.id ?? `${tone}-${index}`;
          return (
            <div
              key={key}
              className={`pointer-events-auto rounded border px-4 py-3 text-sm shadow ${className}`}
            >
              {item.text}
            </div>
          );
        })}
      </div>
    </div>
  );

  if (typeof document === 'undefined') {
    return content;
  }

  return createPortal(content, document.body);
}

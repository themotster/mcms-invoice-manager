export function normalizeProductionItems(items = []) {
  if (!Array.isArray(items)) return [];
  return items
    .map((item, index) => {
      if (!item) return null;
      const id = item.id != null ? String(item.id) : `production-${index}`;
      return {
        id,
        name: item.name != null ? String(item.name) : '',
        description: item.description != null ? String(item.description) : '',
        cost: item.cost != null ? String(item.cost) : '',
        markup: item.markup != null ? String(item.markup) : '',
        notes: item.notes != null ? String(item.notes) : ''
      };
    })
    .filter(Boolean);
}

export function calculateProductionItemTotal(item) {
  if (!item) return 0;
  const costNumber = Number(item.cost);
  const markupNumber = Number(item.markup);
  const base = Number.isFinite(costNumber) ? costNumber : 0;
  const markupFraction = Number.isFinite(markupNumber) ? markupNumber / 100 : 0;
  const total = base + base * markupFraction;
  return Number.isFinite(total) ? total : 0;
}

export function calculateProductionTotal(items = []) {
  const normalized = normalizeProductionItems(items);
  return normalized.reduce((sum, item) => sum + calculateProductionItemTotal(item), 0);
}

export function calculateDiscountValue({ type = 'amount', value, subtotal }) {
  const raw = Number(value);
  const baseSubtotal = Number(subtotal);
  const safeSubtotal = Number.isFinite(baseSubtotal) ? baseSubtotal : 0;
  if (!Number.isFinite(raw) || raw <= 0) return 0;
  if (type === 'percent') {
    return Math.max((safeSubtotal * raw) / 100, 0);
  }
  return Math.max(raw, 0);
}

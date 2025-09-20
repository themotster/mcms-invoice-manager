import {
  normalizeProductionItems,
  calculateProductionItemTotal,
  calculateProductionTotal,
  calculateDiscountValue
} from '../helpers/pricing';

describe('normalizeProductionItems', () => {
  it('returns normalized production entries with string fields and ids', () => {
    const result = normalizeProductionItems([
      { id: 1, name: 'AV Co', cost: 250, markup: 20 },
      { name: 'Lighting', description: 'Wash lights', cost: '100.5', notes: 'Need rigging' }
    ]);

    expect(result).toEqual([
      {
        id: '1',
        name: 'AV Co',
        description: '',
        cost: '250',
        markup: '20',
        notes: ''
      },
      {
        id: 'production-1',
        name: 'Lighting',
        description: 'Wash lights',
        cost: '100.5',
        markup: '',
        notes: 'Need rigging'
      }
    ]);
  });
});

describe('calculateProductionItemTotal', () => {
  it('applies markup percentage to base cost', () => {
    expect(calculateProductionItemTotal({ cost: '100', markup: '10' })).toBeCloseTo(110);
    expect(calculateProductionItemTotal({ cost: '200', markup: 0 })).toBeCloseTo(200);
    expect(calculateProductionItemTotal({ cost: 'invalid', markup: '5' })).toBe(0);
  });
});

describe('calculateProductionTotal', () => {
  it('sums normalized production items', () => {
    const total = calculateProductionTotal([
      { id: 'a', cost: '100', markup: '10' },
      { id: 'b', cost: '50', markup: '0' }
    ]);
    expect(total).toBeCloseTo(160);
  });
});

describe('calculateDiscountValue', () => {
  it('handles amount discounts', () => {
    expect(calculateDiscountValue({ type: 'amount', value: '50', subtotal: 500 })).toBe(50);
  });

  it('handles percent discounts against subtotal', () => {
    expect(calculateDiscountValue({ type: 'percent', value: '10', subtotal: 200 })).toBeCloseTo(20);
  });

  it('returns zero for invalid input', () => {
    expect(calculateDiscountValue({ type: 'percent', value: 'abc', subtotal: 200 })).toBe(0);
    expect(calculateDiscountValue({ type: 'amount', value: '-5', subtotal: 200 })).toBe(0);
  });
});

describe('formatProductionSummary', () => {
  it('formats production summary', () => {
    expect(formatProductionSummary([])).toBe('No production items');
    expect(formatProductionSummary([{ id: 'a' }])).toBe('1 production item');
    expect(formatProductionSummary([{ id: 'a' }, { id: 'b' }])).toBe('2 production items');
  });
});

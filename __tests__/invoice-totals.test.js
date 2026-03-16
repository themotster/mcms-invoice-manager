/**
 * Unit tests for MCMS invoice totals formula (benchmark behaviour).
 * balance_due = max(0, subtotal - discount - received); discount and received stored negative.
 */
describe('Invoice totals (MCMS benchmark)', () => {
  function computeBalanceDue(subtotal, discountAmount, amountReceived) {
    const subtotalVal = Number.isFinite(Number(subtotal)) ? Number(subtotal) : 0;
    const discountVal = Math.abs(Number(discountAmount) || 0);
    const receivedVal = Math.abs(Number(amountReceived) || 0);
    return Math.max(0, subtotalVal - discountVal - receivedVal);
  }

  function discountAsStored(amount) {
    const n = Number(amount);
    if (!Number.isFinite(n) || n === 0) return 0;
    return -Math.abs(n);
  }

  function receivedAsStored(amount) {
    const n = Number(amount);
    if (!Number.isFinite(n) || n === 0) return 0;
    return -Math.abs(n);
  }

  it('balance_due = subtotal - discount - received', () => {
    expect(computeBalanceDue(1525, 100, 200)).toBe(1225);
  });

  it('balance_due with zero discount and received equals subtotal', () => {
    expect(computeBalanceDue(500, 0, 0)).toBe(500);
  });

  it('balance_due is never negative', () => {
    expect(computeBalanceDue(100, 50, 60)).toBe(0);
  });

  it('discount is stored as negative', () => {
    expect(discountAsStored(100)).toBe(-100);
    expect(discountAsStored(0)).toBe(0);
  });

  it('received is stored as negative', () => {
    expect(receivedAsStored(200)).toBe(-200);
    expect(receivedAsStored(0)).toBe(0);
  });
});

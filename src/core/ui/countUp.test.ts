import { countUpValue } from './countUp';

describe('countUpValue (easeOutCubic)', () => {
  it('is pinned at the endpoints', () => {
    expect(countUpValue(0, 100, 0)).toBe(0);
    expect(countUpValue(0, 100, 1)).toBe(100);
  });

  it('clamps out-of-range progress', () => {
    expect(countUpValue(0, 100, -0.5)).toBe(0);
    expect(countUpValue(0, 100, 2)).toBe(100);
  });

  it('eases out — past the linear midpoint by the halfway point', () => {
    // easeOutCubic(0.5) = 1 - 0.5^3 = 0.875 → well ahead of linear 0.5.
    expect(countUpValue(0, 100, 0.5)).toBeCloseTo(87.5, 5);
  });

  it('is monotonic increasing across the run', () => {
    let prev = -Infinity;
    for (let t = 0; t <= 1.0001; t += 0.1) {
      const v = countUpValue(0, 50, t);
      expect(v).toBeGreaterThanOrEqual(prev);
      prev = v;
    }
  });
});

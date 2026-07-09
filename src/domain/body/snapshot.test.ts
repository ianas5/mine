import { BODY_FIELDS, type BodyField } from './fields';
import {
  deriveBmi,
  frequentlyLoggedFields,
  latestFieldValues,
  resolveBmi,
  weightLogWithDeltas,
  type BodySnapshot,
  type BodyValues,
} from './snapshot';

const snap = (date: string, values: Partial<BodyValues> = {}): BodySnapshot => {
  const base = {} as Record<BodyField, number | null>;
  for (const f of BODY_FIELDS) base[f] = null;
  return { date, ...base, ...values };
};

describe('deriveBmi / resolveBmi (FITNESS_DOMAIN §5.2)', () => {
  it('derives weight / heightM²', () => {
    expect(deriveBmi(80, 180)).toBeCloseTo(24.69, 2);
  });

  it('is null without both weight and a positive height', () => {
    expect(deriveBmi(null, 180)).toBeNull();
    expect(deriveBmi(80, null)).toBeNull();
    expect(deriveBmi(80, 0)).toBeNull();
  });

  it('prefers an entered BMI over the derived one', () => {
    expect(resolveBmi(snap('2026-07-09', { bmi: 25, weightKg: 80 }), 180)).toBe(25);
    expect(resolveBmi(snap('2026-07-09', { weightKg: 80 }), 180)).toBeCloseTo(24.69, 2);
  });
});

describe('latestFieldValues (§5.1, newest-first input)', () => {
  it('takes the most recent non-null value per field', () => {
    const latest = latestFieldValues([
      snap('2026-07-09', { weightKg: 80 }),
      snap('2026-07-01', { weightKg: 82, waistCm: 85 }),
    ]);
    expect(latest.weightKg).toEqual({ value: 80, date: '2026-07-09' });
    expect(latest.waistCm).toEqual({ value: 85, date: '2026-07-01' });
    expect(latest.chestCm).toBeNull();
  });
});

describe('frequentlyLoggedFields (UI_UX §5.2 ≥ 50%)', () => {
  it('expands fields present in at least half of sessions', () => {
    const fields = frequentlyLoggedFields([
      snap('2026-07-09', { weightKg: 80, waistCm: 85 }),
      snap('2026-07-02', { weightKg: 81 }),
    ]);
    expect(fields.has('weightKg')).toBe(true); // 2/2
    expect(fields.has('waistCm')).toBe(true); // 1/2 = 0.5
    expect(fields.has('chestCm')).toBe(false); // 0/2
  });

  it('defaults to weight only with no history', () => {
    const fields = frequentlyLoggedFields([]);
    expect([...fields]).toEqual(['weightKg']);
  });
});

describe('weightLogWithDeltas (Measurements home)', () => {
  it('computes the change from the previous (older) weigh-in', () => {
    const log = weightLogWithDeltas([
      { date: '2026-07-09', weightKg: 80 },
      { date: '2026-07-02', weightKg: 82 },
      { date: '2026-06-25', weightKg: 82.5 },
    ]);
    expect(log[0]).toEqual({ date: '2026-07-09', weightKg: 80, deltaKg: -2 });
    expect(log[1]).toEqual({ date: '2026-07-02', weightKg: 82, deltaKg: -0.5 });
    expect(log[2]?.deltaKg).toBeNull(); // oldest has no prior
  });
});

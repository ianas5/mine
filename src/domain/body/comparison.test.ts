import { BODY_FIELDS, type BodyField } from './fields';
import { compareSnapshots, type FieldComparison } from './comparison';
import type { BodySnapshot, BodyValues } from './snapshot';

const snap = (date: string, values: Partial<BodyValues> = {}): BodySnapshot => {
  const base = {} as Record<BodyField, number | null>;
  for (const f of BODY_FIELDS) base[f] = null;
  return { date, ...base, ...values };
};

const get = (result: FieldComparison[], field: BodyField): FieldComparison =>
  result.find((c) => c.field === field)!;

describe('compareSnapshots (FITNESS_DOMAIN §5.4 / §5.3 / §6.4)', () => {
  it('never fabricates a baseline: a field on only one date is incomparable', () => {
    const a = snap('2026-06-01', { weightKg: 82 });
    const b = snap('2026-07-01', { weightKg: 80, waistCm: 85 }); // waist only on B
    const waist = get(compareSnapshots(a, b), 'waistCm');
    expect(waist.direction).toBe('incomparable');
    expect(waist.deltaAbs).toBeNull();
    expect(waist.deltaPct).toBeNull();
    expect(waist.a).toBeNull();
    expect(waist.b).toBe(85);
  });

  it('computes absolute and percentage change for fields present on both', () => {
    const a = snap('2026-06-01', { waistCm: 90 });
    const b = snap('2026-07-01', { waistCm: 85.5 });
    const waist = get(compareSnapshots(a, b), 'waistCm');
    expect(waist.deltaAbs).toBe(-4.5);
    expect(waist.deltaPct).toBe(-5); // -4.5/90*100
    expect(waist.direction).toBe('improving'); // waist lower = better
  });

  it('maps direction by §5.3 (waist down = improving, arm down = declining)', () => {
    const a = snap('a', { waistCm: 90, leftArmCm: 40 });
    const b = snap('b', { waistCm: 92, leftArmCm: 38 });
    const result = compareSnapshots(a, b);
    expect(get(result, 'waistCm').direction).toBe('declining'); // waist up = worse
    expect(get(result, 'leftArmCm').direction).toBe('declining'); // arm down = worse
  });

  it('calls a sub-deadband change stable (§6.4)', () => {
    // waist deadband is 0.5 cm; a 0.3 cm change is stable.
    const stable = get(
      compareSnapshots(snap('a', { waistCm: 85 }), snap('b', { waistCm: 85.3 })),
      'waistCm',
    );
    expect(stable.direction).toBe('stable');
    expect(stable.deltaAbs).toBe(0.3); // the number is still shown
  });

  it('marks weight and BMI changes neutral (no fixed good direction)', () => {
    const weight = get(
      compareSnapshots(snap('a', { weightKg: 82 }), snap('b', { weightKg: 80 })),
      'weightKg',
    );
    expect(weight.direction).toBe('neutral');
    expect(weight.deltaAbs).toBe(-2);
  });

  it('leaves percentage undefined when the baseline is 0 (§5.4)', () => {
    const cmp = get(
      compareSnapshots(snap('a', { visceralFat: 0 }), snap('b', { visceralFat: 3 })),
      'visceralFat',
    );
    expect(cmp.deltaAbs).toBe(3);
    expect(cmp.deltaPct).toBeNull();
  });

  it('derives BMI per date so it is comparable when weight + height are known', () => {
    const a = snap('a', { weightKg: 84 });
    const b = snap('b', { weightKg: 80 });
    const bmi = get(compareSnapshots(a, b, 180), 'bmi');
    expect(bmi.a).toBeCloseTo(25.93, 1);
    expect(bmi.b).toBeCloseTo(24.69, 1);
    expect(bmi.direction).toBe('neutral');
  });
});

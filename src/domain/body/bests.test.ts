import { bestFieldValues, isFieldBest } from './comparison';
import type { BodySnapshot } from './snapshot';

const snap = (over: Partial<BodySnapshot> & { date: string }): BodySnapshot =>
  ({
    weightKg: null,
    bodyFatPct: null,
    muscleMassKg: null,
    visceralFat: null,
    bmi: null,
    neckCm: null,
    chestCm: null,
    waistCm: null,
    hipsCm: null,
    leftArmCm: null,
    rightArmCm: null,
    leftForearmCm: null,
    rightForearmCm: null,
    leftThighCm: null,
    rightThighCm: null,
    leftCalfCm: null,
    rightCalfCm: null,
    ...over,
  }) as BodySnapshot;

describe('bestFieldValues (delight — measurement best)', () => {
  it('takes the lowest for lower-is-better and highest for higher-is-better', () => {
    const best = bestFieldValues([
      snap({ date: '2026-01-01', waistCm: 84, leftArmCm: 38 }),
      snap({ date: '2026-02-01', waistCm: 82, leftArmCm: 39 }),
      snap({ date: '2026-03-01', waistCm: 83, leftArmCm: 37 }),
    ]);
    expect(best.waistCm).toBe(82); // lowest waist
    expect(best.leftArmCm).toBe(39); // highest arm
  });

  it('excludes neutral-direction fields (weight, BMI have no fixed best)', () => {
    const best = bestFieldValues([snap({ date: '2026-01-01', weightKg: 80, bmi: 24 })]);
    expect(best.weightKg).toBeUndefined();
    expect(best.bmi).toBeUndefined();
  });
});

describe('isFieldBest', () => {
  it('is true only when strictly beating the prior best in the good direction', () => {
    expect(isFieldBest('waistCm', 82, 81)).toBe(true); // lower waist beats
    expect(isFieldBest('waistCm', 82, 82)).toBe(false); // tie is not a beat
    expect(isFieldBest('waistCm', 82, 83)).toBe(false); // higher waist is worse
    expect(isFieldBest('leftArmCm', 39, 40)).toBe(true); // bigger arm beats
    expect(isFieldBest('leftArmCm', 39, 38)).toBe(false);
  });

  it('is never a best with no prior best or on a neutral field', () => {
    expect(isFieldBest('waistCm', null, 80)).toBe(false); // nothing to beat
    expect(isFieldBest('weightKg', 90, 70)).toBe(false); // neutral direction
  });
});

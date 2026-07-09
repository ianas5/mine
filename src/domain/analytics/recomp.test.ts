import { BODY_FIELDS, type BodyField, type BodySnapshot } from '@/domain/body';

import { computeRecompSignal } from './recomp';

function snap(date: string, values: Partial<Record<BodyField, number>>): BodySnapshot {
  const base = Object.fromEntries(BODY_FIELDS.map((f) => [f, null])) as Record<
    BodyField,
    number | null
  >;
  return { date, ...base, ...values };
}

const today = '2026-03-15';

describe('computeRecompSignal (FITNESS_DOMAIN §6.5)', () => {
  it('fires when weight is stable and a fat-down/muscle-up marker holds', () => {
    const r = computeRecompSignal(
      [
        snap('2026-01-25', { weightKg: 80, waistCm: 84, bodyFatPct: 18 }),
        snap('2026-03-15', { weightKg: 80.4, waistCm: 82.5, bodyFatPct: 17 }), // stable wt, waist −1.5
      ],
      today,
    );
    expect(r.status).toBe('ok');
    if (r.status !== 'ok') return;
    expect(r.value.fired).toBe(true);
    expect(r.value.markers.length).toBeGreaterThan(0);
  });

  it('does NOT fire when weight dropped meaningfully (that is fat loss, not recomp)', () => {
    const r = computeRecompSignal(
      [
        snap('2026-01-25', { weightKg: 84, waistCm: 86 }),
        snap('2026-03-15', { weightKg: 80, waistCm: 83 }), // −4 kg → not stable
      ],
      today,
    );
    if (r.status === 'ok') expect(r.value.fired).toBe(false);
  });

  it('does not fire when weight is stable but no marker moved enough', () => {
    const r = computeRecompSignal(
      [
        snap('2026-01-25', { weightKg: 80, waistCm: 84 }),
        snap('2026-03-15', { weightKg: 80.2, waistCm: 83.8 }), // waist only −0.2 cm
      ],
      today,
    );
    if (r.status === 'ok') expect(r.value.fired).toBe(false);
  });

  it('is insufficient-data below 4 weeks of span', () => {
    const r = computeRecompSignal(
      [
        snap('2026-03-01', { weightKg: 80, waistCm: 84 }),
        snap('2026-03-15', { weightKg: 80, waistCm: 82 }), // only 14 days
      ],
      today,
    );
    expect(r.status).toBe('insufficient-data');
  });
});

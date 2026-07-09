import { dayAdherence, remainingMacros, type TargetMacros } from './adherence';

const target: TargetMacros = { kcal: 2000, proteinG: 160, carbG: 200, fatG: 60, waterMl: 3000 };

describe('dayAdherence (FITNESS_DOMAIN §4.3)', () => {
  it('treats protein as a floor with a 90% near-band, never over', () => {
    expect(
      dayAdherence(target, { kcal: 2000, proteinG: 160, carbG: 200, fatG: 60 }, 3000).protein,
    ).toBe('hit');
    expect(
      dayAdherence(target, { kcal: 2000, proteinG: 150, carbG: 200, fatG: 60 }, 3000).protein,
    ).toBe('near'); // 93.75%
    expect(
      dayAdherence(target, { kcal: 2000, proteinG: 120, carbG: 200, fatG: 60 }, 3000).protein,
    ).toBe('under');
    // Way over protein is still just a hit (more is fine).
    expect(
      dayAdherence(target, { kcal: 2000, proteinG: 300, carbG: 200, fatG: 60 }, 3000).protein,
    ).toBe('hit');
  });

  it('bands calories at ±10% and carbs/fat at ±15%', () => {
    const a = dayAdherence(target, { kcal: 2250, proteinG: 160, carbG: 235, fatG: 51 }, 3000);
    expect(a.calories).toBe('over'); // 2250 > 2200
    expect(a.carbs).toBe('over'); // 235 > 230
    expect(a.fat).toBe('hit'); // 51 within [51,69]
    const b = dayAdherence(target, { kcal: 1750, proteinG: 160, carbG: 165, fatG: 50 }, 3000);
    expect(b.calories).toBe('under'); // 1750 < 1800
    expect(b.carbs).toBe('under'); // 165 < 170
    expect(b.fat).toBe('under'); // 50 < 51
  });

  it('judges water only when the target sets a goal', () => {
    expect(dayAdherence(target, { kcal: 0, proteinG: 0, carbG: 0, fatG: 0 }, 3000).water).toBe(
      'hit',
    );
    expect(dayAdherence(target, { kcal: 0, proteinG: 0, carbG: 0, fatG: 0 }, 1000).water).toBe(
      'under',
    );
    expect(
      dayAdherence({ ...target, waterMl: null }, { kcal: 0, proteinG: 0, carbG: 0, fatG: 0 }, 1000)
        .water,
    ).toBeNull();
  });
});

describe('remainingMacros (§4.2, negative = over)', () => {
  it('subtracts consumed from target and allows negatives', () => {
    expect(remainingMacros(target, { kcal: 1500, proteinG: 120, carbG: 250, fatG: 40 })).toEqual({
      kcal: 500,
      proteinG: 40,
      carbG: -50,
      fatG: 20,
    });
  });
});

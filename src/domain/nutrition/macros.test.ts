import { atwaterKcal, isKcalImplausible, scalePortion, sumMacros } from './macros';

describe('scalePortion (FITNESS_DOMAIN §4.2)', () => {
  const food = { servingAmount: 100, kcal: 165, proteinG: 31, carbG: 0, fatG: 3.6 };

  it('scales macros by loggedAmount / servingAmount', () => {
    expect(scalePortion(food, 150)).toEqual({ kcal: 248, proteinG: 46.5, carbG: 0, fatG: 5.4 });
  });

  it('rounds kcal to an integer and grams to 0.1 g', () => {
    expect(scalePortion(food, 50)).toEqual({ kcal: 83, proteinG: 15.5, carbG: 0, fatG: 1.8 });
  });

  it('scales to zero rather than dividing by a zero serving', () => {
    expect(scalePortion({ ...food, servingAmount: 0 }, 100)).toEqual({
      kcal: 0,
      proteinG: 0,
      carbG: 0,
      fatG: 0,
    });
  });
});

describe('sumMacros (FITNESS_DOMAIN §4.2 day totals)', () => {
  it('sums entries and re-rounds grams to 0.1 g', () => {
    expect(
      sumMacros([
        { kcal: 248, proteinG: 46.5, carbG: 0, fatG: 5.4 },
        { kcal: 200, proteinG: 10.1, carbG: 40.2, fatG: 2.3 },
      ]),
    ).toEqual({ kcal: 448, proteinG: 56.6, carbG: 40.2, fatG: 7.7 });
  });

  it('is zero for an empty day', () => {
    expect(sumMacros([])).toEqual({ kcal: 0, proteinG: 0, carbG: 0, fatG: 0 });
  });
});

describe('macro/energy cross-check (FITNESS_DOMAIN §4.2, validation aid only)', () => {
  it('computes the Atwater estimate', () => {
    expect(atwaterKcal({ kcal: 0, proteinG: 31, carbG: 0, fatG: 3.6 })).toBeCloseTo(156.4, 4);
  });

  it('flags a food whose kcal diverges from its macros beyond tolerance', () => {
    // entered 400 kcal, atwater ≈ 156 → implausible.
    expect(isKcalImplausible({ kcal: 400, proteinG: 31, carbG: 0, fatG: 3.6 })).toBe(true);
    // entered 165 kcal, atwater ≈ 156 → within 20 %.
    expect(isKcalImplausible({ kcal: 165, proteinG: 31, carbG: 0, fatG: 3.6 })).toBe(false);
  });

  it('never flags a zero-kcal food (nothing to compare)', () => {
    expect(isKcalImplausible({ kcal: 0, proteinG: 0, carbG: 0, fatG: 0 })).toBe(false);
  });
});

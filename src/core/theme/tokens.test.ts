import { darkColors, darkTheme, lightColors, lightTheme } from './tokens';

describe('theme tokens', () => {
  it('ships dark and light together with identical token keys (DESIGN_SYSTEM rule 6)', () => {
    expect(Object.keys(lightColors).sort()).toEqual(Object.keys(darkColors).sort());
  });

  it('defines a non-empty value for every color token in both themes', () => {
    for (const colors of [darkColors, lightColors]) {
      for (const [token, value] of Object.entries(colors)) {
        expect({ token, value }).toEqual({ token, value: expect.stringMatching(/^(#|rgba\()/) });
      }
    }
  });

  it('shares scale tokens (space, radius, type, motion) across themes', () => {
    expect(darkTheme.space).toBe(lightTheme.space);
    expect(darkTheme.radius).toBe(lightTheme.radius);
    expect(darkTheme.type).toBe(lightTheme.type);
    expect(darkTheme.motion).toBe(lightTheme.motion);
  });

  it('marks each theme with its mode', () => {
    expect(darkTheme.mode).toBe('dark');
    expect(lightTheme.mode).toBe('light');
  });
});

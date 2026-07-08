// Jest environment setup — mocks for native-touching modules.

// react-native-mmkv is a JSI module; tests use an in-memory stand-in.
jest.mock('react-native-mmkv', () => {
  class MMKV {
    constructor() {
      this.map = new Map();
    }
    getString(key) {
      return this.map.get(key);
    }
    set(key, value) {
      this.map.set(key, value);
    }
    delete(key) {
      this.map.delete(key);
    }
  }
  return { MMKV };
});

// expo-haptics touches native modules; tests assert through the token wrapper instead.
jest.mock('expo-haptics', () => ({
  impactAsync: jest.fn(() => Promise.resolve()),
  notificationAsync: jest.fn(() => Promise.resolve()),
  ImpactFeedbackStyle: { Light: 'light' },
  NotificationFeedbackType: { Success: 'success', Warning: 'warning' },
}));

// expo-router navigation is provided by the native runtime; stub it for tests.
jest.mock('expo-router', () => ({
  useRouter: () => ({ push: jest.fn(), back: jest.fn(), replace: jest.fn() }),
  Redirect: () => null,
}));

// expo-crypto is native; use Node's randomUUID in tests.
jest.mock('expo-crypto', () => ({
  randomUUID: () => require('node:crypto').randomUUID(),
}));

// Safe-area insets need a native provider; the library ships a test mock.
jest.mock(
  'react-native-safe-area-context',
  () => require('react-native-safe-area-context/jest/mock').default,
);

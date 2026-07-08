import * as Haptics from 'expo-haptics';
import { Platform } from 'react-native';

/**
 * Haptic tokens — DESIGN_SYSTEM §4: `light` (set logged, chip select, stepper
 * tick), `success` (workout saved, PR), `warning` (destructive confirm).
 * All haptics go through this token map so intensity policy stays central.
 */
export type HapticToken = 'light' | 'success' | 'warning';

export function triggerHaptic(token: HapticToken): void {
  if (Platform.OS === 'web') {
    return;
  }
  switch (token) {
    case 'light':
      void Haptics.impactAsync(Haptics.ImpactFeedbackStyle.Light).catch(() => undefined);
      break;
    case 'success':
      void Haptics.notificationAsync(Haptics.NotificationFeedbackType.Success).catch(
        () => undefined,
      );
      break;
    case 'warning':
      void Haptics.notificationAsync(Haptics.NotificationFeedbackType.Warning).catch(
        () => undefined,
      );
      break;
  }
}

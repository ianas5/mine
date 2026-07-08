import { useEffect, useRef, useState, type ReactNode } from 'react';
// (hideTimer stays a ref: it is only touched inside effects/handlers.)
import { Animated, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { triggerHaptic, useTheme } from '@/core/theme';

type ToastTone = 'neutral' | 'success';

interface ToastMessage {
  readonly text: string;
  readonly tone: ToastTone;
}

const DISPLAY_MS = 2500;

let listener: ((message: ToastMessage) => void) | null = null;

/** Show a single-line toast (one at a time, 2.5s). Success tone adds the success haptic. */
export function showToast(text: string, tone: ToastTone = 'neutral'): void {
  if (tone === 'success') {
    triggerHaptic('success');
  }
  listener?.({ text, tone });
}

/** Mounted once at the app root; renders the active toast above the tab bar. */
export function ToastHost(): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const [message, setMessage] = useState<ToastMessage | null>(null);
  const [opacity] = useState(() => new Animated.Value(0));
  const hideTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    listener = (next) => {
      if (hideTimer.current !== null) {
        clearTimeout(hideTimer.current);
      }
      setMessage(next);
      Animated.timing(opacity, {
        toValue: 1,
        duration: theme.motion.fast,
        useNativeDriver: true,
      }).start();
      hideTimer.current = setTimeout(() => {
        Animated.timing(opacity, {
          toValue: 0,
          duration: theme.motion.base,
          useNativeDriver: true,
        }).start(() => setMessage(null));
      }, DISPLAY_MS);
    };
    return () => {
      listener = null;
      if (hideTimer.current !== null) {
        clearTimeout(hideTimer.current);
      }
    };
  }, [opacity, theme.motion.base, theme.motion.fast]);

  if (message === null) {
    return null;
  }
  return (
    <View
      pointerEvents="none"
      style={{
        position: 'absolute',
        left: theme.space.lg,
        right: theme.space.lg,
        bottom: insets.bottom + 72,
        alignItems: 'center',
      }}
    >
      <Animated.View
        accessibilityRole="alert"
        style={{
          opacity,
          backgroundColor: theme.color.surfaceRaised,
          borderColor: message.tone === 'success' ? theme.color.positive : theme.color.border,
          borderWidth: 1,
          borderRadius: theme.radius.md,
          paddingHorizontal: theme.space.lg,
          paddingVertical: theme.space.md,
        }}
      >
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
          {message.text}
        </Text>
      </Animated.View>
    </View>
  );
}

import { useEffect, useRef, useState, type ReactNode } from 'react';
// (hideTimer stays a ref: it is only touched inside effects/handlers.)
import { Animated, Pressable, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { triggerHaptic, useTheme } from '@/core/theme';

type ToastTone = 'neutral' | 'success';

/** An optional single action (e.g. Undo) shown on the toast (UI_UX §6). */
export interface ToastAction {
  readonly label: string;
  readonly onPress: () => void;
}

interface ToastMessage {
  readonly text: string;
  readonly tone: ToastTone;
  readonly action: ToastAction | undefined;
}

const DISPLAY_MS = 2500;
const ACTION_DISPLAY_MS = 5000; // Undo toasts linger (UI_UX §6)

let listener: ((message: ToastMessage) => void) | null = null;

/**
 * Show a single-line toast (one at a time). Success tone adds the success haptic.
 * An optional `action` (e.g. Undo) makes the toast tappable and lingers 5 s.
 */
export function showToast(text: string, tone: ToastTone = 'neutral', action?: ToastAction): void {
  if (tone === 'success') {
    triggerHaptic('success');
  }
  listener?.({ text, tone, action });
}

/** Mounted once at the app root; renders the active toast above the tab bar. */
export function ToastHost(): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const [message, setMessage] = useState<ToastMessage | null>(null);
  const [opacity] = useState(() => new Animated.Value(0));
  const hideTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const dismiss = (): void => {
    if (hideTimer.current !== null) {
      clearTimeout(hideTimer.current);
      hideTimer.current = null;
    }
    Animated.timing(opacity, {
      toValue: 0,
      duration: theme.motion.base,
      useNativeDriver: true,
    }).start(() => setMessage(null));
  };

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
      hideTimer.current = setTimeout(
        () => {
          Animated.timing(opacity, {
            toValue: 0,
            duration: theme.motion.base,
            useNativeDriver: true,
          }).start(() => setMessage(null));
        },
        next.action ? ACTION_DISPLAY_MS : DISPLAY_MS,
      );
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

  const action = message.action;
  return (
    <View
      pointerEvents="box-none"
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
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.md,
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
        {action ? (
          <Pressable
            onPress={() => {
              action.onPress();
              dismiss();
            }}
            accessibilityRole="button"
            accessibilityLabel={action.label}
            hitSlop={theme.space.sm}
          >
            <Text style={{ ...theme.type.bodyStrong, color: theme.color.accent }}>
              {action.label}
            </Text>
          </Pressable>
        ) : null}
      </Animated.View>
    </View>
  );
}
